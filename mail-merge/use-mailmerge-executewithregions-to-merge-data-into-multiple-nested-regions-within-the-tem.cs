using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title.
        builder.Writeln("Customer Orders Report");
        builder.Writeln();

        // ---- Begin outer region: Orders ----
        builder.InsertField("MERGEFIELD TableStart:Orders");
        builder.Writeln();

        // Order fields.
        builder.InsertField("MERGEFIELD OrderID");
        builder.Write("\t");
        builder.InsertField("MERGEFIELD OrderDate");
        builder.Writeln();

        // ---- Begin inner region: OrderDetails ----
        builder.InsertField("MERGEFIELD TableStart:OrderDetails");
        builder.Writeln();

        // Detail fields.
        builder.InsertField("MERGEFIELD ProductName");
        builder.Write("\t");
        builder.InsertField("MERGEFIELD Quantity");
        builder.Writeln();

        // End inner region.
        builder.InsertField("MERGEFIELD TableEnd:OrderDetails");
        builder.Writeln();

        // End outer region.
        builder.InsertField("MERGEFIELD TableEnd:Orders");
        builder.Writeln();

        // Prepare data tables.
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("OrderDate", typeof(DateTime));

        DataTable details = new DataTable("OrderDetails");
        details.Columns.Add("OrderID", typeof(int));
        details.Columns.Add("ProductName", typeof(string));
        details.Columns.Add("Quantity", typeof(int));

        // Sample data for orders.
        orders.Rows.Add(1, new DateTime(2023, 1, 15));
        orders.Rows.Add(2, new DateTime(2023, 2, 20));

        // Sample data for order details.
        details.Rows.Add(1, "Apple", 10);
        details.Rows.Add(1, "Banana", 5);
        details.Rows.Add(2, "Orange", 7);
        details.Rows.Add(2, "Grapes", 12);

        // Combine tables into a DataSet and define relation for nested merge.
        DataSet dataSet = new DataSet();
        dataSet.Tables.Add(orders);
        dataSet.Tables.Add(details);
        dataSet.Relations.Add("Order_Details", orders.Columns["OrderID"], details.Columns["OrderID"]);

        // Execute mail merge with regions.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Remove any unused merge fields.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields;

        // Save the merged document.
        string outputPath = "MergedDocument.docx";
        doc.Save(outputPath);

        // Validation: ensure file exists and contains expected content.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");

        FileInfo fileInfo = new FileInfo(outputPath);
        if (fileInfo.Length == 0)
            throw new InvalidOperationException("The merged document is empty.");

        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();
        if (!resultText.Contains("Apple") || !resultText.Contains("Orange"))
            throw new InvalidOperationException("Merged content does not contain expected values.");
    }
}
