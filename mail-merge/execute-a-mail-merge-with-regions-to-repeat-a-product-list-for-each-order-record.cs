using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class MailMergeWithRegionsExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the template with an outer region for Orders.
        builder.InsertField("MERGEFIELD TableStart:Orders");
        builder.Writeln();

        builder.Writeln("Order Report");
        builder.Writeln();

        builder.Writeln("Order ID: ");
        builder.InsertField("MERGEFIELD OrderID");
        builder.Writeln();

        builder.Writeln("Customer: ");
        builder.InsertField("MERGEFIELD CustomerName");
        builder.Writeln();

        // Insert inner region for product list.
        builder.InsertField("MERGEFIELD TableStart:Products");
        builder.StartTable();

        // Header row.
        builder.InsertCell(); builder.Write("Product");
        builder.InsertCell(); builder.Write("Quantity");
        builder.EndRow();

        // Data row with merge fields.
        builder.InsertCell(); builder.InsertField("MERGEFIELD ProductName");
        builder.InsertCell(); builder.InsertField("MERGEFIELD Quantity");
        builder.EndRow();

        builder.EndTable();
        builder.InsertField("MERGEFIELD TableEnd:Products");
        builder.Writeln();

        builder.InsertField("MERGEFIELD TableEnd:Orders");
        builder.Writeln();

        // Prepare data source.
        DataSet dataSet = new DataSet();

        // Orders table (main table).
        DataTable ordersTable = new DataTable("Orders");
        ordersTable.Columns.Add("OrderID", typeof(int));
        ordersTable.Columns.Add("CustomerName", typeof(string));
        ordersTable.Rows.Add(1, "Alice");
        ordersTable.Rows.Add(2, "Bob");
        dataSet.Tables.Add(ordersTable);

        // Products table (region table).
        DataTable productsTable = new DataTable("Products");
        productsTable.Columns.Add("OrderID", typeof(int)); // Link to Orders.OrderID
        productsTable.Columns.Add("ProductName", typeof(string));
        productsTable.Columns.Add("Quantity", typeof(int));
        // Products for Order 1
        productsTable.Rows.Add(1, "Apple", 5);
        productsTable.Rows.Add(1, "Banana", 3);
        // Products for Order 2
        productsTable.Rows.Add(2, "Orange", 7);
        productsTable.Rows.Add(2, "Grape", 2);
        dataSet.Tables.Add(productsTable);

        // Execute mail merge with regions.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save merged document.
        string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Validate that the merged document contains expected data.
        Document mergedDoc = new Document(outputPath);
        string mergedText = mergedDoc.GetText();

        if (string.IsNullOrWhiteSpace(mergedText) ||
            !mergedText.Contains("Alice") ||
            !mergedText.Contains("Bob") ||
            !mergedText.Contains("Apple") ||
            !mergedText.Contains("Orange"))
        {
            throw new InvalidOperationException("Mail merge failed: merged content is missing expected values.");
        }

        Console.WriteLine("Mail merge completed successfully. Output saved to " + Path.GetFullPath(outputPath));
    }
}
