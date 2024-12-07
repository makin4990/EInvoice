using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace EInvoice.Services;

public class CustomerInvoiceConverter
{
    private XNamespace ns;
    private XNamespace cbc;
    private XNamespace cac;

    public async Task<MemoryStream> ConvertToExcel(List<IFormFile> files)
    {

        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = SetupWorksheet(workbook, "Invoice Data");
            var invoiceDtos = new List<InvoiceDto>();

            ns = "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2";
            cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
            cbc = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
            int currentRow = 2;

            foreach (var file in files)
            {
                string xmlContent;
                using (var reader = new StreamReader(file.OpenReadStream()))
                {
                    xmlContent = await reader.ReadToEndAsync();
                }

                XDocument doc = XDocument.Parse(xmlContent);
                WriteFileToWorksheetLines(doc, worksheet, ref currentRow, invoiceDtos);
            }
            var worksheet2 = SetupWorksheet(workbook, "GroupedData");

            GroupAndSummarizeWorksheet(worksheet2, invoiceDtos);

            var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;

            return stream;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error converting e-invoice to Excel: {ex.Message}", ex);
        }
    }
    public async Task<MemoryStream> ConvertToExcel(List<IFormFile> files, Dictionary<string, decimal> currencyInfo)
    {

        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = CreateExcelWithHeader(workbook, "Invoice Data");
            var invoiceDtos = new List<InvoiceDto>();

            ns = "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2";
            cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
            cbc = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";

            int currentRow = 3;

            foreach (var file in files)
            {
                string xmlContent;
                using (var reader = new StreamReader(file.OpenReadStream()))
                {
                    xmlContent = await reader.ReadToEndAsync();
                }

                XDocument doc = XDocument.Parse(xmlContent);
                WriteFileToWorksheetLines(doc, worksheet, ref currentRow, invoiceDtos,currencyInfo);
            }

            var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;

            return stream;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error converting e-invoice to Excel: {ex.Message}", ex);
        }
    }
    private IXLWorksheet SetupWorksheet(XLWorkbook workbook, string tabName)
    {
        var worksheet = workbook.Worksheets.Add(tabName);

        // Add headers
        worksheet.Cell(1, 1).Value = "Sıra No";
        worksheet.Cell(1, 2).Value = "Alış Faturasının Tarihi";
        worksheet.Cell(1, 3).Value = "Alış Faturasının Serisi";
        worksheet.Cell(1, 4).Value = "Alış Faturasının Sıra No'su";
        worksheet.Cell(1, 5).Value = "Satıcının Adı-Soyadı / Ünvanı";
        worksheet.Cell(1, 6).Value = "Satıcının Vergi Kimlik Numarası / TC Kimlik Numarası";
        worksheet.Cell(1, 7).Value = "Alınan Mal ve/veya Hizmetin Cinsi";
        worksheet.Cell(1, 8).Value = "Alınan Mal ve/veya Hizmetin Miktarı";
        worksheet.Cell(1, 9).Value = "Alınan Mal ve/veya Hizmetin KDV Hariç Tutarı";
        worksheet.Cell(1, 10).Value = "KDV'si";
        worksheet.Cell(1, 11).Value = "";
        worksheet.Cell(1, 12).Value = "";
        worksheet.Cell(1, 13).Value = "";
        worksheet.Cell(1, 14).Value = "GGB Tescil No'su (Alış İthalat İse)";
        worksheet.Cell(1, 15).Value = "Belgenin İndirim Hakkının Kullanıldığı KDV Dönemi";

        // Format header row
        var headerRow = worksheet.Row(1);
        headerRow.Style.Font.Bold = true;
        headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;
        return worksheet;
    }
    private void WriteFileToWorksheetLines(XDocument doc, IXLWorksheet worksheet, ref int currentRow, List<InvoiceDto> invoiceDtos)
    {
        var invoices = doc.Descendants(ns + "Invoice").ToList();
        foreach (var invoice in invoices)
        {
            // Extract basic invoice information
            string invoiceDate = invoice.Elements(cbc + "IssueDate").FirstOrDefault()?.Value ?? "";
            string invoiceNumber = invoice.Elements(cbc + "ID").FirstOrDefault()?.Value ?? "";
            string serviceNo = string.Empty;
            // Extract supplier information

            #region supplierName

            var accountingSupplierParty = invoice.Elements(cac + "AccountingSupplierParty").FirstOrDefault();
            var supplierPartyName = accountingSupplierParty?.Descendants(cac + "PartyName").FirstOrDefault();
            string supplierName = supplierPartyName?.Descendants(cbc + "Name")?.FirstOrDefault()?.Value ?? string.Empty;

            #endregion

            #region customerName

            var accountingCustomerParty = invoice.Elements(cac + "AccountingCustomerParty").FirstOrDefault();
            var customerParty = accountingCustomerParty?.Descendants(cac + "PartyName").FirstOrDefault();
            string customerName = customerParty?.Descendants(cbc + "Name")?.FirstOrDefault()?.Value ?? string.Empty;

            #endregion

            #region supplierIdentity

            var supplierParty = accountingSupplierParty?.Descendants(cac + "Party").FirstOrDefault();
            var supplierPartyIdentification = supplierParty?.Elements(cac + "PartyIdentification")?.FirstOrDefault();
            string supplierIdentity = supplierPartyIdentification?.Elements(cbc + "ID")?.FirstOrDefault()?.Value ?? string.Empty;

            #endregion

            // Extract customer information
            // Extract line items
            var invoiceLines = invoice.Descendants(cac + "InvoiceLine");
            foreach (var line in invoiceLines)
            {
                string itemSpec = line.Elements(cac + "Item")
                    .Elements(cbc + "Name")
                    .FirstOrDefault()?.Value ?? "";

                decimal quantity = decimal.Parse(line.Elements(cbc + "InvoicedQuantity")
                    .FirstOrDefault()?.Value ?? "0");

                var taxTotal = line.Descendants(cac + "TaxTotal");
                decimal taxAmount = taxTotal.Descendants(cac + "TaxSubtotal")
                     .Elements(cbc + "TaxAmount")
                     .Select(e => (decimal)e)
                     .FirstOrDefault();

                var taxableAmount = taxTotal
                                   .Descendants(cac + "TaxSubtotal")
                                   .Elements(cbc + "TaxableAmount")
                                   .Select(e => (decimal)e)
                                   .FirstOrDefault();

                worksheet.Cell(currentRow, 1).Value = currentRow - 1;
                worksheet.Cell(currentRow, 2).Value = invoiceDate;
                worksheet.Cell(currentRow, 3).Value = serviceNo;
                worksheet.Cell(currentRow, 4).Value = invoiceNumber;
                worksheet.Cell(currentRow, 5).Value = supplierName;
                worksheet.Cell(currentRow, 6).Value = supplierIdentity;
                worksheet.Cell(currentRow, 7).Value = itemSpec;
                worksheet.Cell(currentRow, 8).Value = $"{quantity:F2} ADET";
                worksheet.Cell(currentRow, 9).Value = taxableAmount;
                worksheet.Cell(currentRow, 10).Value = taxAmount;
                worksheet.Cell(currentRow, 11).Value = string.Empty;
                worksheet.Cell(currentRow, 12).Value = string.Empty;
                worksheet.Cell(currentRow, 13).Value = string.Empty;
                worksheet.Cell(currentRow, 14).Value = string.Empty;
                worksheet.Cell(currentRow, 15).Value = string.Empty;
                invoiceDtos.Add(new()
                {
                    InvoiceDate = invoiceDate,
                    InvoiceOrderNo = invoiceNumber,
                    SellerName = supplierName,
                    SellerIdentity = supplierIdentity,
                    ProductType = itemSpec,
                    Quantity = quantity,
                    VatFreeAmount = taxAmount,
                    VatAmount = taxableAmount

                });


                currentRow++;
            }
        }

        // Auto-fit columns
        worksheet.Columns().AdjustToContents();

        // Add some basic formatting
        var dataRange = worksheet.Range(1, 1, currentRow - 1, 15);
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Format number columns
        worksheet.Column(6).Style.NumberFormat.NumberFormatId = 1; // Number format for Quantity
        worksheet.Column(7).Style.NumberFormat.Format = "#,##0.00"; // Currency format for Unit Price
        worksheet.Column(8).Style.NumberFormat.Format = "#,##0.00"; // Currency format for Total Amount
        worksheet.Column(9).Style.NumberFormat.Format = "#,##0.00"; // Currency format for Tax Amount

    }
    private void WriteFileToWorksheetLines(XDocument doc, IXLWorksheet worksheet, ref int currentRow, List<InvoiceDto> invoiceDtos, Dictionary<string, decimal> currencyInfo)
    {
        var invoices = doc.Descendants(ns + "Invoice").ToList();
        foreach (var invoice in invoices)
        {
            string invoiceDate = invoice.Elements(cbc + "IssueDate").FirstOrDefault()?.Value ?? "";
            string invoiceNumber = invoice.Elements(cbc + "ID").FirstOrDefault()?.Value ?? "";
            string serviceNo = string.Empty;

            #region supplierName

            var accountingSupplierParty = invoice.Elements(cac + "AccountingSupplierParty").FirstOrDefault();
            var supplierPartyName = accountingSupplierParty?.Descendants(cac + "PartyName").FirstOrDefault();
            string supplierName = supplierPartyName?.Descendants(cbc + "Name")?.FirstOrDefault()?.Value ?? string.Empty;

            #endregion

            #region customerName

            var accountingCustomerParty = invoice.Elements(cac + "AccountingCustomerParty").FirstOrDefault();
            var customerParty = accountingCustomerParty?.Descendants(cac + "PartyName").FirstOrDefault();
            string customerName = customerParty?.Descendants(cbc + "Name")?.FirstOrDefault()?.Value ?? string.Empty;

            #endregion

            #region supplierIdentity

            var supplierParty = accountingSupplierParty?.Descendants(cac + "Party").FirstOrDefault();
            var supplierPartyIdentification = supplierParty?.Elements(cac + "PartyIdentification")?.FirstOrDefault();
            string supplierIdentity = supplierPartyIdentification?.Elements(cbc + "ID")?.FirstOrDefault()?.Value ?? string.Empty;

            #endregion

            var invoiceLines = invoice.Descendants(cac + "InvoiceLine");
            foreach (var line in invoiceLines)
            {
                string itemSpec = line.Elements(cac + "Item")
                    .Elements(cbc + "Name")
                    .FirstOrDefault()?.Value ?? "";

                decimal quantity = decimal.Parse(line.Elements(cbc + "InvoicedQuantity")
                    .FirstOrDefault()?.Value ?? "0");

                var taxTotal = line.Descendants(cac + "TaxTotal");
                var taxableAmount = taxTotal
                    .Descendants(cac + "TaxSubtotal")
                    .Elements(cbc + "TaxableAmount")
                    .Select(e => (decimal)e)
                    .FirstOrDefault();
                //Percent
                var taxAmountElement = taxTotal.Elements(cbc + "TaxAmount").FirstOrDefault();
                var currencyId = taxAmountElement?.Attribute("currencyID").Value;

                string  vatPercentageValue = taxTotal.Descendants(cac + "TaxSubtotal")?
                                                     .Elements(cbc + "Percent")?
                                                     .FirstOrDefault()?.Value ?? "0";
                vatPercentageValue = vatPercentageValue.Replace('.', ',');
                decimal vatPercentage = decimal.Parse(vatPercentageValue);
                decimal taxAmount = taxTotal.Descendants(cac + "TaxSubtotal")
                    .Elements(cbc + "TaxAmount")
                    .Select(e => (decimal)e)
                    .FirstOrDefault();

                worksheet.Cell(currentRow, 1).Value = currentRow - 1;
                worksheet.Cell(currentRow, 2).Value = invoiceDate;
                worksheet.Cell(currentRow, 3).Value = serviceNo;
                worksheet.Cell(currentRow, 4).Value = invoiceNumber;
                worksheet.Cell(currentRow, 5).Value = supplierName;
                worksheet.Cell(currentRow, 6).Value = supplierIdentity;
                worksheet.Cell(currentRow, 7).Value = itemSpec;
                worksheet.Cell(currentRow, 8).Value = $"{quantity:F2} ADET";
                worksheet.Cell(currentRow, 9).Value = taxableAmount;
                worksheet.Cell(currentRow, 10).Value = vatPercentage;
                worksheet.Cell(currentRow, 11).Value = taxAmount;
                worksheet.Cell(currentRow, 12).Value = string.Empty;
                worksheet.Cell(currentRow, 13).Value = string.Empty;
                worksheet.Cell(currentRow, 14).Value = string.Empty;
                worksheet.Cell(currentRow, 15).Value = string.Empty;
                invoiceDtos.Add(new()
                {
                    InvoiceDate = invoiceDate,
                    InvoiceOrderNo = invoiceNumber,
                    SellerName = supplierName,
                    SellerIdentity = supplierIdentity,
                    ProductType = itemSpec,
                    Quantity = quantity,
                    VatFreeAmount = taxableAmount,
                    VatAmount = taxAmount,
                    CurrencyValue = currencyInfo[currencyId]
                });


                currentRow++;
            }

        }
        worksheet.Cell(currentRow, 8).Value = "Toplam";
        worksheet.Cell(currentRow, 9).Value = invoiceDtos.Sum(i=> i.VatFreeAmount*i.CurrencyValue);
        worksheet.Cell(currentRow, 10).Value = "0.00";
        worksheet.Cell(currentRow, 11).Value = invoiceDtos.Sum(i => i.VatAmount * i.CurrencyValue);


        worksheet.Columns().AdjustToContents();

        var dataRange = worksheet.Range(1, 1, currentRow - 1, 18);
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        worksheet.Column(6).Style.NumberFormat.NumberFormatId = 1; // Number format for Quantity
        worksheet.Column(7).Style.NumberFormat.Format = "#,##0.00"; // Currency format for Unit Price
        worksheet.Column(8).Style.NumberFormat.Format = "#,##0.00"; // Currency format for Total Amount
        worksheet.Column(9).Style.NumberFormat.Format = "#,##0.00"; // Currency format for Tax Amount

    }
    private void GroupAndSummarizeWorksheet(IXLWorksheet worksheet, List<InvoiceDto> invoiceDtos)
    {

        var groupedData = invoiceDtos
            .GroupBy(row => row.InvoiceOrderNo)
            .Select(g => new InvoiceDto
            {
                InvoiceDate = g.First().InvoiceDate ?? string.Empty,
                InvoiceOrderNo = g.Key,
                SellerName = g.First().SellerName,
                SellerIdentity = g.First().SellerIdentity,
                ProductType = string.Join(',', g.Select(i => i.ProductType)),
                Quantity = g.Sum(row => row.Quantity),
                VatFreeAmount = g.Sum(row => row.VatFreeAmount),
                VatAmount = g.Sum(row => row.VatAmount),
            });


        int currentRow = 2;
        foreach (var group in groupedData)
        {
            worksheet.Cell(currentRow, 1).Value = currentRow - 1;
            worksheet.Cell(currentRow, 2).Value = group.InvoiceDate;
            worksheet.Cell(currentRow, 3).Value = "";
            worksheet.Cell(currentRow, 4).Value = group.InvoiceOrderNo;
            worksheet.Cell(currentRow, 5).Value = group.SellerName;
            worksheet.Cell(currentRow, 6).Value = group.SellerIdentity;
            worksheet.Cell(currentRow, 7).Value = group.ProductType;
            worksheet.Cell(currentRow, 8).Value = $"{group.Quantity:F2} ADET";
            worksheet.Cell(currentRow, 9).Value = group.VatFreeAmount;
            worksheet.Cell(currentRow, 10).Value = group.VatAmount;
            worksheet.Cell(currentRow, 11).Value = string.Empty;
            worksheet.Cell(currentRow, 12).Value = string.Empty;
            worksheet.Cell(currentRow, 13).Value = string.Empty;
            worksheet.Cell(currentRow, 14).Value = string.Empty;
            worksheet.Cell(currentRow, 15).Value = string.Empty;
            currentRow++;
        }

        worksheet.Columns().AdjustToContents();
        var dataRange = worksheet.Range(1, 1, currentRow - 1, 15);
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        worksheet.Column(6).Style.NumberFormat.NumberFormatId = 1; // Number format for Quantity
        worksheet.Column(7).Style.NumberFormat.Format = "#,##0.00"; // Currency format for Unit Price
        worksheet.Column(8).Style.NumberFormat.Format = "#,##0.00"; // Currency format for Total Amount
        worksheet.Column(9).Style.NumberFormat.Format = "#,##0.00"; // Currency format for Tax Amount
    }
    private IXLWorksheet CreateExcelWithHeader(XLWorkbook workbook, string tabName)
    {
        var worksheet = workbook.AddWorksheet("Sheet1");
        var headers = new Dictionary<char, string> {
                {'A',"Sıra No"},
                {'B',"Satış Faturasının Tarihi"},
                {'C',"Satış Faturasının Serisi"},
                {'D',"Satış Faturasının Sıra No'su"},
                {'E',"Alıcının Adı-Soyadı / Ünvanı"},
                {'F',"Alıcının Vergi Kimlik Numarası / TC Kimlik Numarası"},
                {'G',"Satılan Mal ve/veya Hizmetin Cinsi"},
                {'H',"Satılan Mal ve/veya Hizmetin Miktarı"},
                {'I',"Satılan Mal ve/veya Hizmetin KDV Hariç Tutarı"},
                {'J',"Kdv Oranı (%)"},
                {'K',"KDV'si"},
                {'L',"İade işlem Türü"},
                {'M',"GÇB Tescil No (405 İşlem Kodundan iadelerde)"},
                {'P',"Sektör Bilgisi Numarası"},
                {'Q',"Alt Sektör Bilgisi Numarası"},
                {'R', "Konut Teslimi Yapılan Kişi/Kurum"}
            };
        // A to M: Each column has its own header spanning two rows
        for (char col = 'A'; col <= 'M'; col++)
        {
            string header = headers[col];
            var cell = worksheet.Cell(1, col - 'A' + 1);
            cell.Value = header;
            cell.Style.Font.Bold = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
            // Merge the two rows
            worksheet.Range(1, col - 'A' + 1, 2, col - 'A' + 1).Merge();
        }
 
        // N and O: Merged header with individual sub-headers
        worksheet.Cell(1, 14).Value = "Satışı Yapılan Taşınmaza Ait";
        worksheet.Cell(1, 14).Style.Font.Bold = true;
        worksheet.Cell(2, 14).Style.Alignment.WrapText = true;
        worksheet.Cell(1, 14).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        worksheet.Cell(1, 14).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        worksheet.Range(1, 14, 1, 15).Merge(); // Merge N1:O1

        worksheet.Cell(2, 14).Value = "Zemin Sistem No"; // N2
        worksheet.Cell(2, 14).Style.Alignment.WrapText = true;

        worksheet.Cell(2, 15).Value = "Tapu Kayıt Yevmiye Numarası"; // O2
        worksheet.Cell(2, 15).Style.Alignment.WrapText = true;

        worksheet.Cell(2, 14).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        worksheet.Cell(2, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        // P to R: Same pattern as A to M
        for (char col = 'P'; col <= 'R'; col++)
        {
            string header = headers[col];
            var cell = worksheet.Cell(1, col - 'A' + 1);
            cell.Value = header;
            cell.Style.Font.Bold = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;
            // Merge the two rows
            worksheet.Range(1, col - 'A' + 1, 2, col - 'A' + 1).Merge();

        }

        // Apply borders for clarity
        worksheet.Range(1, 1, 2, 18).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        worksheet.Range(1, 1, 2, 18).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Adjust column widths for better visibility
        // Adjust column widths
        SetColWitdh(worksheet);
        var headerRow1 = worksheet.Row(1);
        headerRow1.Style.Font.Bold = true;
        headerRow1.Style.Fill.BackgroundColor = XLColor.LightGray;
        var headerRow2 = worksheet.Row(2);
        headerRow2.Style.Font.Bold = true;
        headerRow2.Style.Fill.BackgroundColor = XLColor.LightGray;
        return worksheet;
    }

    private void SetColWitdh(IXLWorksheet worksheet)
    {
        worksheet.Column(1).Style.Alignment.WrapText = false;
        worksheet.Column(10).Style.Alignment.WrapText = false;
        worksheet.Column(11).Style.Alignment.WrapText = false;

        worksheet.Column(1).Width = 5.57;
        for (int i = 1; i<=18; i++)
            worksheet.Column(i).Width = 10.14;
        worksheet.Column(13).Width = 16.14;
    } 
}
