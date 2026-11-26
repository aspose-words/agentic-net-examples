using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert simple (non‑region) merge fields.
        builder.Writeln("Order Number: ");
        builder.InsertField("MERGEFIELD OrderNumber");
        builder.Writeln();

        builder.Writeln("Customer: ");
        builder.InsertField("MERGEFIELD CustomerName");
        builder.Writeln();

        // Insert a mail‑merge region for order items.
        // Start of the region.
        builder.InsertField("MERGEFIELD TableStart:OrderItems");

        // Begin a table that will be repeated for each row in the region.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data row with merge fields.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD ItemName");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Quantity");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // End of the region.
        builder.InsertField("MERGEFIELD TableEnd:OrderItems");

        // (Optional) Save the template for inspection.
        const string templatePath = "OrderTemplate.docx";
        doc.Save(templatePath);

        // Prepare data for the simple merge fields.
        string[] fieldNames = { "OrderNumber", "CustomerName" };
        object[] fieldValues = { "12345", "John Doe" };

        // Prepare a DataTable for the OrderItems region.
        DataTable itemsTable = new DataTable("OrderItems");
        itemsTable.Columns.Add("ItemName", typeof(string));
        itemsTable.Columns.Add("Quantity", typeof(int));

        itemsTable.Rows.Add("Apple", 3);
        itemsTable.Rows.Add("Banana", 5);
        itemsTable.Rows.Add("Cherry", 2);

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);
        doc.MailMerge.ExecuteWithRegions(itemsTable);

        // Save the merged document.
        const string outputPath = "MergedOrder.docx";
        doc.Save(outputPath);

        // Validate that key values appear in the merged output.
        string mergedText = doc.GetText();
        if (!mergedText.Contains("12345") ||
            !mergedText.Contains("John Doe") ||
            !mergedText.Contains("Apple") ||
            !mergedText.Contains("Banana") ||
            !mergedText.Contains("Cherry"))
        {
            throw new InvalidOperationException("Merged document does not contain expected data.");
        }

        // Indicate successful completion.
        Console.WriteLine($"Mail merge completed successfully. Output saved to '{outputPath}'.");
    }
}
