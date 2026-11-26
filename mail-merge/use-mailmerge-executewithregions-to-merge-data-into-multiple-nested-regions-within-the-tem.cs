using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a title.
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Start the outer region "Orders".
        builder.InsertField("MERGEFIELD TableStart:Orders", null);
        builder.Writeln();

        // Insert fields for each order.
        builder.Writeln("Order ID: ");
        builder.InsertField("MERGEFIELD OrderID", null);
        builder.Writeln();
        builder.Writeln("Customer: ");
        builder.InsertField("MERGEFIELD CustomerName", null);
        builder.Writeln();

        // Start the inner region "OrderDetails".
        builder.InsertField("MERGEFIELD TableStart:OrderDetails", null);
        builder.Writeln();

        // Insert a simple table for order details.
        builder.StartTable();
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Product", null);
        builder.InsertCell();
        builder.InsertField("MERGEFIELD Quantity", null);
        builder.EndRow();
        builder.EndTable();

        // End the inner region.
        builder.InsertField("MERGEFIELD TableEnd:OrderDetails", null);
        builder.Writeln();

        // End the outer region.
        builder.InsertField("MERGEFIELD TableEnd:Orders", null);
        builder.Writeln();

        // Prepare data tables.
        DataTable ordersTable = new DataTable("Orders");
        ordersTable.Columns.Add("OrderID", typeof(int));
        ordersTable.Columns.Add("CustomerName", typeof(string));

        ordersTable.Rows.Add(1, "Alice");
        ordersTable.Rows.Add(2, "Bob");

        DataTable detailsTable = new DataTable("OrderDetails");
        detailsTable.Columns.Add("OrderID", typeof(int)); // Foreign key to Orders
        detailsTable.Columns.Add("Product", typeof(string));
        detailsTable.Columns.Add("Quantity", typeof(int));

        detailsTable.Rows.Add(1, "Apples", 5);
        detailsTable.Rows.Add(1, "Bananas", 3);
        detailsTable.Rows.Add(2, "Oranges", 7);
        detailsTable.Rows.Add(2, "Grapes", 2);

        // Create a DataSet and add tables with a relation.
        DataSet dataSet = new DataSet();
        dataSet.Tables.Add(ordersTable);
        dataSet.Tables.Add(detailsTable);
        dataSet.Relations.Add("Order_OrderDetails",
            ordersTable.Columns["OrderID"],
            detailsTable.Columns["OrderID"]);

        // Clean up any unused fields after merge.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

        // Execute mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.docx");
        doc.Save(outputPath);

        // Validate that the output contains expected merged data.
        string mergedText = doc.GetText();
        if (string.IsNullOrWhiteSpace(mergedText) ||
            !mergedText.Contains("Alice") ||
            !mergedText.Contains("Bob") ||
            !mergedText.Contains("Apples") ||
            !mergedText.Contains("Oranges"))
        {
            throw new InvalidOperationException("Mail merge did not produce the expected output.");
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("Mail merge completed successfully. Output saved to: " + outputPath);
    }
}
