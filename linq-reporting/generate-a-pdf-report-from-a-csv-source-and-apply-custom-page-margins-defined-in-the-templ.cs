using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing on .NET Core.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File names used in the working directory.
        const string templatePath = "ReportTemplate.docx";
        const string csvPath = "Data.csv";
        const string outputPath = "Report.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with headers and a few rows of data.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Product,Quantity,Price",
            "Apple,10,0.5",
            "Banana,5,0.3",
            "Orange,8,0.6"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // ---------------------------------------------------------------
        // 2. Build the Word template programmatically and insert LINQ tags.
        // ---------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Apply custom page margins (1 inch = 72 points).
        builder.PageSetup.TopMargin = 72;
        builder.PageSetup.BottomMargin = 72;
        builder.PageSetup.LeftMargin = 72;
        builder.PageSetup.RightMargin = 72;

        // Title.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Product Report");
        builder.Writeln(); // Empty line.

        // Begin foreach loop over CSV rows (data source name will be "data").
        builder.Writeln("<<foreach [row in data]>>");

        // Table with header row.
        Table table = builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Font.Bold = true;
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row – values are taken from the current CSV row.
        builder.InsertCell();
        builder.Font.Bold = false;
        builder.Writeln("<<[row.Product]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Quantity]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Price]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template and bind the CSV data source.
        // ---------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "data");

        // ---------------------------------------------------------------
        // 4. Save the populated document as PDF.
        // ---------------------------------------------------------------
        reportDoc.Save(outputPath, SaveFormat.Pdf);
    }
}
