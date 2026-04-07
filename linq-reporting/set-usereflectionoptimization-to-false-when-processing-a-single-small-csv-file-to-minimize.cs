using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "people.csv");
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a small CSV file with headers and a couple of rows.
        // -----------------------------------------------------------------
        // The CSV contains two columns: Name and Age.
        string[] csvLines =
        {
            "Name,Age",
            "Alice,30",
            "Bob,25"
        };
        File.WriteAllLines(csvPath, csvLines);

        // -----------------------------------------------------------------
        // 2. Build a simple Word template that uses LINQ Reporting tags.
        // -----------------------------------------------------------------
        // The template iterates over the CSV rows and writes each person's data.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach block over the data source named "persons".
        builder.Writeln("<<foreach [p in persons]>>");
        // Inside the loop write the fields.
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template back for reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Disable reflection optimization for this small CSV scenario.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = false;

        // -----------------------------------------------------------------
        // 5. Prepare the CSV data source with header support.
        // -----------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true); // true => first line has headers
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 6. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root name "persons" must match the name used in the template tags.
        engine.BuildReport(reportDoc, dataSource, "persons");

        // -----------------------------------------------------------------
        // 7. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
