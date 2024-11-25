using System.Xml.Linq;
using ClosedXML.Excel;

namespace EInvoice.Services;

public class CustomerInvoiceConverter
{
        private XNamespace ns;
    private XNamespace cbc;
    private XNamespace cac;

    public async Task<MemoryStream> ConvertToExcel(List<IFormFile> files)
    {
        // Read XML content

        try
        {
            // Create a new Excel workbook
            using var workbook = new XLWorkbook();
            var worksheet = SetupWorksheet(workbook);

            // Assuming namespace for e-invoice (modify according to your XML structure)
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

                // Load the XML file
                XDocument doc = XDocument.Parse(xmlContent);
                // Extract invoice data
                WriteFileToWorksheetLines(doc,worksheet, ref currentRow);
            }


            // Save the Excel file
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

    private void WriteFileToWorksheetLines(XDocument doc, IXLWorksheet worksheet, ref int  currentRow)
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
                var taxAmount = taxTotal
                    .Elements(cbc + "TaxAmount")
                    .Select(e => (decimal)e)
                    .FirstOrDefault();

                decimal taxableAmount = taxTotal.Descendants(cac + "TaxSubtotal")
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
                ;
                worksheet.Cell(currentRow, 9).Value = taxAmount;
                worksheet.Cell(currentRow, 10).Value = taxableAmount;
                worksheet.Cell(currentRow, 11).Value = string.Empty;
                worksheet.Cell(currentRow, 12).Value = string.Empty;
                worksheet.Cell(currentRow, 13).Value = string.Empty;
                worksheet.Cell(currentRow, 14).Value = string.Empty;
                worksheet.Cell(currentRow, 15).Value = string.Empty;

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

    private IXLWorksheet SetupWorksheet(XLWorkbook workbook)
    {
        var worksheet = workbook.Worksheets.Add("Invoice Data");

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
}