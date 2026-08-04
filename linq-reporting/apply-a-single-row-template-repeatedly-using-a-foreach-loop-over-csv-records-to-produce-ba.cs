using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for CSV parsing on some platforms).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the working directory.
        const string templatePath = "Template.docx";
        const string csvPath = "people.csv";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a simple CSV file with headers and sample data.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Name,Age",
            "Alice,30",
            "Bob,25",
            "Charlie,35"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // Step 2: Build the template document programmatically.
        // The template contains a foreach block that iterates over the CSV rows.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 3: Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Configure CSV loading options – the first line contains headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);

        // Create a CSV data source based on the file and options.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // Step 4: Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // Step 5: Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
