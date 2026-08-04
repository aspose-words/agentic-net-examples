using System;
using System.IO;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Path for the template document.
        string templatePath = Path.Combine(outputDir, "template.docx");

        // Create a DOCX template with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [row in ds.Products]>>");
        builder.Writeln("Name: <<[row.Name]>>, Price: $<<[row.Price]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template document.
        Document reportDoc = new Document(templatePath);

        // Create a DataSet with a sample DataTable.
        DataSet ds = new DataSet();
        DataTable productsTable = new DataTable("Products");
        productsTable.Columns.Add("Name", typeof(string));
        productsTable.Columns.Add("Price", typeof(decimal));
        productsTable.Rows.Add("Apple", 0.5m);
        productsTable.Rows.Add("Banana", 0.3m);
        productsTable.Rows.Add("Cherry", 1.2m);
        ds.Tables.Add(productsTable);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, ds, "ds");

        // Save the generated report as PDF.
        string pdfPath = Path.Combine(outputDir, "report.pdf");
        reportDoc.Save(pdfPath, SaveFormat.Pdf);
    }
}
