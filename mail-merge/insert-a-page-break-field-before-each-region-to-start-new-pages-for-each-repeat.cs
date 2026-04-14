using System;
using System.Data;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // First mail merge region: "Customers"
        // Insert a page break before the region so each repeat starts on a new page.
        // -------------------------------------------------
        builder.InsertBreak(BreakType.PageBreak);
        // Begin the region.
        builder.InsertField("MERGEFIELD TableStart:Customers");
        // Content that will be repeated for each customer.
        builder.Write("Customer: ");
        builder.InsertField("MERGEFIELD Name");
        builder.Writeln();
        // End the region.
        builder.InsertField("MERGEFIELD TableEnd:Customers");
        builder.Writeln();

        // -------------------------------------------------
        // Second mail merge region: "Orders"
        // Insert a page break before this region as well.
        // -------------------------------------------------
        builder.InsertBreak(BreakType.PageBreak);
        builder.InsertField("MERGEFIELD TableStart:Orders");
        builder.Write("Order: ");
        builder.InsertField("MERGEFIELD Item");
        builder.Write(", Qty: ");
        builder.InsertField("MERGEFIELD Quantity");
        // End the region.
        builder.InsertField("MERGEFIELD TableEnd:Orders");
        builder.Writeln();

        // ------------------------------
        // Prepare data for the "Customers" region.
        // ------------------------------
        DataTable customers = new DataTable("Customers");
        customers.Columns.Add("Name");
        customers.Rows.Add("Alice");
        customers.Rows.Add("Bob");

        // ------------------------------
        // Prepare data for the "Orders" region.
        // ------------------------------
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("Item");
        orders.Columns.Add("Quantity");
        orders.Rows.Add("Apples", "10");
        orders.Rows.Add("Bananas", "5");
        orders.Rows.Add("Oranges", "8");

        // Perform mail merge for the first region.
        doc.MailMerge.ExecuteWithRegions(customers);

        // Perform mail merge for the second region.
        doc.MailMerge.ExecuteWithRegions(orders);

        // Save the resulting document.
        doc.Save("MailMergeWithPageBreaks.docx");
    }
}
