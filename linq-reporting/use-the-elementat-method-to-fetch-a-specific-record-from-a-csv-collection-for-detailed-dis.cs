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

        // File paths (relative to the working directory).
        string csvPath = "people.csv";
        string templatePath = "template.docx";
        string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with headers and a few records.
        // -----------------------------------------------------------------
        string[] csvLines =
        {
            "Id,Name,Age",
            "1,John Doe,30",
            "2,Jane Smith,25",
            "3,Bob Johnson,40"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a Word template that uses LINQ Reporting tags.
        //    The template will display the second record (index 1) using
        //    the ElementAt method.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("All records (foreach):");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("- Id: <<[p.Id]>>, Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        builder.Writeln();
        builder.Writeln("Detailed view of the second record (ElementAt):");
        // ElementAt is zero‑based; 1 fetches the second row.
        builder.Writeln("Id: <<[persons.ElementAt(1).Id]>>");
        builder.Writeln("Name: <<[persons.ElementAt(1).Name]>>");
        builder.Writeln("Age: <<[persons.ElementAt(1).Age]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and prepare the CSV data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // CSV options: first line contains headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
