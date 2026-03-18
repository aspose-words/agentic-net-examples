using System;
using System.Collections.Generic;

public class Invoice
{
    public int Id { get; set; }
    public decimal Amount { get; set; }

    public override string ToString() => $"Invoice Id={Id}, Amount={Amount}";
}

public class Program
{
    public static void Main()
    {
        // Create a List<Invoice> with sample data
        List<Invoice> invoices = new List<Invoice>
        {
            new Invoice { Id = 1, Amount = 150.00m },
            new Invoice { Id = 2, Amount = 299.99m },
            new Invoice { Id = 3, Amount = 75.50m }
        };

        // Iterate over the collection using foreach with an explicit generic type
        foreach (Invoice invoice in invoices)
        {
            // Example processing: output each invoice
            Console.WriteLine(invoice);
        }
    }
}
