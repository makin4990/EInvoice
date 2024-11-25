using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;

namespace EInvoice.Data;

public class EInvoiceDbConztext : IdentityDbContext
{
    public EInvoiceDbConztext(DbContextOptions<EInvoiceDbConztext> options)
        : base(options)
    {
    }
}