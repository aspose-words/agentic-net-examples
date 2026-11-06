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

        // Begin the Orders region.
        builder.InsertField("MERGEFIELD TableStart:Orders", null);
        builder.Writeln(); // Move to next line.

        // Insert order-level merge fields.
        builder.InsertField("MERGEFIELD OrderID", null);
        builder.Write("\t");
        builder.InsertField("MERGEFIELD CustomerName", null);
        builder.Writeln();

        // Begin the nested Products region inside the Orders region.
        builder.InsertField("MERGEFIELD TableStart:Products", null);
        builder.Writeln();

        // Insert product-level merge fields.
        builder.InsertField("MERGEFIELD ProductName", null);
        builder.Write("\t");
        builder.InsertField("MERGEFIELD Quantity", null);
        builder.Writeln();

        // End the Products region.
        builder.InsertField("MERGEFIELD TableEnd:Products", null);
        builder.Writeln();

        // End the Orders region.
        builder.InsertField("MERGEFIELD TableEnd:Orders", null);
        builder.Writeln();

        // Prepare the data source.
        DataTable ordersTable = new DataTable("Orders");
        ordersTable.Columns.Add("OrderID", typeof(int));
        ordersTable.Columns.Add("CustomerName", typeof(string));

        DataTable productsTable = new DataTable("Products");
        productsTable.Columns.Add("OrderID", typeof(int));
        productsTable.Columns.Add("ProductName", typeof(string));
        productsTable.Columns.Add("Quantity", typeof(int));

        // Populate orders.
        ordersTable.Rows.Add(1, "Alice");
        ordersTable.Rows.Add(2, "Bob");

        // Populate products for each order.
        productsTable.Rows.Add(1, "Apple", 5);
        productsTable.Rows.Add(1, "Banana", 3);
        productsTable.Rows.Add(2, "Carrot", 7);
        productsTable.Rows.Add(2, "Dates", 2);
        productsTable.Rows.Add(2, "Eggplant", 4);

        // Create a DataSet and add tables.
        DataSet dataSet = new DataSet();
        dataSet.Tables.Add(ordersTable);
        dataSet.Tables.Add(productsTable);

        // Define relation between Orders and Products on OrderID.
        dataSet.Relations.Add("Order_Products",
            ordersTable.Columns["OrderID"],
            productsTable.Columns["OrderID"]);

        // Execute mail merge with regions using the DataSet.
        doc.MailMerge.ExecuteWithRegions(dataSet);

        // Save the merged document.
        string outputPath = "MergedOutput.docx";
        doc.Save(outputPath);

        // Validation: ensure the output file exists and is not empty.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged output file was not created.");

        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The merged output file is empty.");

        // Optionally, you could add further validation of content here.
    }
}
