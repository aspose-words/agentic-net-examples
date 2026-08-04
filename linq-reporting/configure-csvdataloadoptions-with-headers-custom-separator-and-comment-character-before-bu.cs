using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting; // CsvDataLoadOptions, CsvDataSource

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible non‑UTF8 CSV files.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Define file paths.
        string workDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(workDir, "Data");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        string csvPath = Path.Combine(dataDir, "people.csv");
        string templatePath = Path.Combine(dataDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with headers, custom delimiter ';' and comment character '$'.
        // -----------------------------------------------------------------
        // The file contains a comment line (starts with $) that will be ignored.
        string[] csvLines =
        {
            "$ This is a comment line and will be ignored by the parser",
            "Name;Age;City",
            "Alice;30;New York",
            "Bob;25;London",
            "Charlie;35;Paris"
        };
        File.WriteAllLines(csvPath, csvLines);

        // -----------------------------------------------------------------
        // 2. Build a template document that uses LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>, City: <<[person.City]>>");
        builder.Writeln("<</foreach>>");

        // Save the template so it can be loaded later (demonstrates load/save lifecycle).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Configure CsvDataLoadOptions: headers present, ';' delimiter, '$' comment char.
        // -----------------------------------------------------------------
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.Delimiter = ';';
        loadOptions.CommentChar = '$';
        // QuoteChar can stay default (") – not required for this data.

        // -----------------------------------------------------------------
        // 5. Create the CSV data source with the configured options.
        // -----------------------------------------------------------------
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // -----------------------------------------------------------------
        // 6. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "persons");

        // -----------------------------------------------------------------
        // 7. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(resultPath);

        // The example finishes without waiting for user input.
    }
}
