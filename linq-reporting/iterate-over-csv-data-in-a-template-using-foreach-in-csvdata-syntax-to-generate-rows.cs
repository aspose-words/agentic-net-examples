using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string templatePath = "Template.docx";
        string csvPath = "Data.csv";
        string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a simple CSV file with headers and sample rows.
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
        // Step 2: Build a Word template that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("People List:");
        builder.Writeln();

        // Begin the foreach loop over the CSV data source named "csvData".
        // Correct syntax: <<foreach [row in csvData]>>
        builder.Writeln("<<foreach [row in csvData]>>");
        // Inside the loop output the fields from each CSV row.
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 3: Load the template and bind the CSV data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Configure CSV loading to treat the first line as headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // Build the report. The data source name used in the template tags is "csvData".
        engine.BuildReport(doc, csvDataSource, "csvData");

        // -----------------------------------------------------------------
        // Step 4: Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
