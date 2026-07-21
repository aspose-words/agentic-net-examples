using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample CSV data.
        string csvPath = "people.csv";
        string[] csvLines =
        {
            "Name,Age",
            "Alice,30",
            "Bob,25",
            "Charlie,35"
        };
        File.WriteAllLines(csvPath, csvLines);

        // Create a LINQ Reporting template.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use a foreach block to iterate over the CSV rows.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Configure CSV data source with headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // Disable reflection optimization for this report.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report using the CSV data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "persons");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
