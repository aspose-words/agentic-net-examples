using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template, CSV data, and the generated report.
        string templatePath = "Template.docx";
        string csvPath = "Data.csv";
        string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple Word template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the CSV rows (named "persons").
        builder.Writeln("<<foreach [person in persons]>>");
        // Output each column value.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("City: <<[person.City]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a CSV file with sample data.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Name,Age,City",
            "Alice,30,New York",
            "Bob,25,London",
            "Charlie,35,Sydney"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 3. Load the template and prepare the CSV data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Configure CSV loading options (first line contains headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        // Use default delimiter ','; other options can be set if needed.

        // Create the CSV data source.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Disable reflection optimization for this report.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = false;

        // -----------------------------------------------------------------
        // 5. Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name used in the template tags is "persons".
        engine.BuildReport(doc, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
