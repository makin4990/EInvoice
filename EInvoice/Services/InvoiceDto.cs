namespace EInvoice.Services;

public class InvoiceDto
{
    public string InvoiceDate { get; set; }
    public string InvoiceOrderNo { get; set; }
    public string SellerName {get; set; }
    public string SellerIdentity {get; set; }
    public string ProductType { get; set; }
    public decimal Quantity {get; set; }
    public decimal VatFreeAmount {get; set; }
    public decimal VatAmount {get; set; }
    public string GGBApproveNo { get; set; }
    public string TaxDiscountDate { get; set; }
    public decimal CurrencyValue { get; set; }

}
