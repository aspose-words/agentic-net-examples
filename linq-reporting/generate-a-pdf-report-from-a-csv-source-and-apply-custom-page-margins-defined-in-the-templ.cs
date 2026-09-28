using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string workingDir = Directory.GetCurrentDirectory();
        string csvPath = Path.Combine(workingDir, "Data.csv");
        string templatePath = Path.Combine(workingDir, "Template.docx");
        string outputPdfPath = Path.Combine(workingDir, "Report.pdf");

        // Create sample CSV data.
        File.WriteAllText(csvPath,
            "Name,Age,City\n" +
            "Alice,30,New York\n" +
            "Bob,25,Los Angeles\n" +
            "Charlie,35,Chicago");

        // -----------------------------------------------------------------
        // Create the template document with custom page margins (1 inch = 72 points).
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.PageSetup.TopMargin = 72;
        builder.PageSetup.BottomMargin = 72;
        builder.PageSetup.LeftMargin = 72;
        builder.PageSetup.RightMargin = 72;

        // Insert LINQ Reporting tags to iterate over CSV rows.
        builder.Writeln("<<foreach [row in data]>>");
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>, City: <<[row.City]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and prepare the CSV data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        var csvOptions = new CsvDataLoadOptions
        {
            HasHeaders = true
            // The default separator is a comma, which matches our CSV data.
        };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, csvData, "data");

        // Save the final report as PDF.
        reportDoc.Save(outputPdfPath);
    }
}
