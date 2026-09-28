using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Paths for the template and the final PDF.
        string templatePath = "template.docx";
        string outputPdfPath = "report.pdf";

        // -------------------------------------------------
        // 1. Create a DOCX template with LINQ Reporting tags.
        // -------------------------------------------------
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln();

        // Begin a foreach loop over the DataTable named "Products".
        builder.Writeln("<<foreach [row in Products]>>");
        // Write each product's name and price.
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Price: $<<[row.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back from file.
        // -------------------------------------------------
        Document doc = new(templatePath);

        // -------------------------------------------------
        // 3. Prepare a DataSet with sample data.
        // -------------------------------------------------
        DataSet dataSet = new();
        DataTable productsTable = new("Products");
        productsTable.Columns.Add("Name", typeof(string));
        productsTable.Columns.Add("Price", typeof(decimal));

        // Add sample rows.
        productsTable.Rows.Add("Apple", 0.99m);
        productsTable.Rows.Add("Banana", 0.59m);
        productsTable.Rows.Add("Cherry", 2.49m);

        dataSet.Tables.Add(productsTable);

        // -------------------------------------------------
        // 4. Build the report using ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new();
        // No special options required for this simple example.
        engine.BuildReport(doc, dataSet, "DataSet");

        // -------------------------------------------------
        // 5. Save the generated report as PDF.
        // -------------------------------------------------
        doc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
