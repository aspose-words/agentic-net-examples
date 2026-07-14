using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholder text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello {{Name}}!");
        builder.Writeln("Your order number is {{OrderId}}.");
        builder.Writeln("Thank you for shopping with us.");

        // Simulate retrieving dynamic data from a database using a DataTable.
        DataTable dataTable = GetSampleData();

        // Assume the first row contains the data we need.
        DataRow row = dataTable.Rows[0];
        string name = row["Name"].ToString();
        string orderId = row["OrderId"].ToString();

        // Replace placeholders in the whole document range.
        // The placeholders are case‑insensitive by default.
        doc.Range.Replace("{{Name}}", name);
        doc.Range.Replace("{{OrderId}}", orderId);

        // Save the resulting document.
        doc.Save("Output.docx");
    }

    // Creates a DataTable that mimics data retrieved from a database.
    private static DataTable GetSampleData()
    {
        DataTable table = new DataTable("CustomerOrders");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("OrderId", typeof(string));

        // Insert a sample record.
        table.Rows.Add("John Doe", "A12345");

        return table;
    }
}
