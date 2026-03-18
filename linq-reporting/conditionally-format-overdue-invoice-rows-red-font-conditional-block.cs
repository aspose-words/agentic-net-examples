using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class ConditionalOverdueInvoices
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample invoice data.
        var invoices = new List<Invoice>
        {
            new Invoice { Number = "INV-001", DueDate = DateTime.Today.AddDays(-5), Amount = 1500.00 },
            new Invoice { Number = "INV-002", DueDate = DateTime.Today.AddDays(3),  Amount =  750.00 },
            new Invoice { Number = "INV-003", DueDate = DateTime.Today.AddDays(-2), Amount =  300.00 },
            new Invoice { Number = "INV-004", DueDate = DateTime.Today.AddDays(10), Amount = 1200.00 }
        };

        // Start a table and add a header row.
        builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Font.Bold = true;
        builder.Write("Invoice #");
        builder.InsertCell();
        builder.Write("Due Date");
        builder.InsertCell();
        builder.Write("Amount");
        builder.EndRow();

        // Reset header formatting.
        builder.Font.Bold = false;

        // Iterate through invoices and apply conditional formatting.
        foreach (var inv in invoices)
        {
            // Determine if the invoice is overdue.
            bool isOverdue = inv.DueDate < DateTime.Today;

            // If overdue, set the font color to red for the entire row.
            if (isOverdue)
                builder.Font.Color = Color.Red;
            else
                builder.Font.Color = Color.Black; // Default color.

            // Invoice number cell.
            builder.InsertCell();
            builder.Write(inv.Number);

            // Due date cell.
            builder.InsertCell();
            builder.Write(inv.DueDate.ToString("d"));

            // Amount cell.
            builder.InsertCell();
            builder.Write(inv.Amount.ToString("C"));

            // End the current row.
            builder.EndRow();

            // Reset font color for the next iteration.
            builder.Font.Color = Color.Black;
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save("ConditionalOverdueInvoices.docx");
    }

    // Simple POCO to hold invoice information.
    class Invoice
    {
        public string Number { get; set; }
        public DateTime DueDate { get; set; }
        public double Amount { get; set; }
    }
}
