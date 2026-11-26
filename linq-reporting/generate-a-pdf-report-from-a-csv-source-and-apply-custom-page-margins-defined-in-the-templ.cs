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
        // Register code page provider required for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with headers and a few rows of data.
        // -----------------------------------------------------------------
        string csvPath = "data.csv";
        string csvContent =
            "Name,Age,Country\n" +
            "Alice,30,USA\n" +
            "Bob,25,Canada\n" +
            "Charlie,35,UK";
        File.WriteAllText(csvPath, csvContent);

        // -----------------------------------------------------------------
        // 2. Build the Word template programmatically.
        //    The template contains LINQ Reporting tags that will be bound
        //    to the CSV data source.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Set custom page margins (1 inch = 72 points).
        builder.PageSetup.LeftMargin = 72;
        builder.PageSetup.RightMargin = 72;
        builder.PageSetup.TopMargin = 72;
        builder.PageSetup.BottomMargin = 72;

        // Title.
        builder.Writeln("Customer Report");
        builder.Writeln();

        // Begin foreach block.
        builder.Writeln("<<foreach [row in rows]>>");

        // Start the table that will be repeated for each row.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("Country");
        builder.EndRow();

        // Data row – each cell contains a field expression.
        builder.InsertCell();
        builder.Writeln("<<[row.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[row.Country]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End of foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the CSV data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options – the file has a header row.
        CsvDataLoadOptions csvOptions = new CsvDataLoadOptions { HasHeaders = true };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvData, "rows");

        // -----------------------------------------------------------------
        // 4. Save the generated report as PDF.
        // -----------------------------------------------------------------
        string outputPdf = "report.pdf";
        reportDoc.Save(outputPdf, SaveFormat.Pdf);
    }
}
