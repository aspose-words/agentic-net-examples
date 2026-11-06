using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class MailMergeRegionExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple merge field for the order number.
        builder.Writeln("Order Number:");
        builder.InsertField("MERGEFIELD OrderNumber");
        builder.Writeln();

        // Insert the start of the region.
        builder.InsertField("MERGEFIELD TableStart:OrderItems");
        builder.Writeln();

        // Build a table that will be repeated for each row in the region.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Data row with merge fields.
        builder.InsertCell();
        builder.InsertField("MERGEFIELD ItemName");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Quantity");
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Price");
        builder.EndRow();

        builder.EndTable();

        // Insert the end of the region.
        builder.InsertField("MERGEFIELD TableEnd:OrderItems");
        builder.Writeln();

        // Prepare simple merge data.
        string[] simpleFieldNames = { "OrderNumber" };
        object[] simpleFieldValues = { "ORD-1001" };

        // Prepare region data in a DataTable.
        DataTable itemsTable = new DataTable("OrderItems");
        itemsTable.Columns.Add("ItemName", typeof(string));
        itemsTable.Columns.Add("Quantity", typeof(int));
        itemsTable.Columns.Add("Price", typeof(decimal));

        itemsTable.Rows.Add("Widget", 4, 19.99m);
        itemsTable.Rows.Add("Gadget", 2, 34.50m);
        itemsTable.Rows.Add("Doohickey", 1, 99.95m);

        // Execute the mail merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;
        doc.MailMerge.Execute(simpleFieldNames, simpleFieldValues);
        doc.MailMerge.ExecuteWithRegions(itemsTable);

        // Save the merged document.
        string outputPath = "MergedOrder.docx";
        doc.Save(outputPath);

        // Validate that the merge produced expected content.
        string mergedText = doc.GetText();
        if (!mergedText.Contains("ORD-1001") ||
            !mergedText.Contains("Widget") ||
            !mergedText.Contains("Gadget") ||
            !mergedText.Contains("Doohickey"))
        {
            throw new InvalidOperationException("Mail merge did not produce the expected output.");
        }

        Console.WriteLine($"Mail merge completed successfully. Output saved to '{outputPath}'.");
    }
}
