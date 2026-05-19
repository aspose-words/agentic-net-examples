using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a simple foreach loop that will iterate over the CSV rows.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Country: <<[person.Country]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Create a CSV file with headers, a custom separator and a comment character.
        string csvPath = Path.Combine(outputDir, "Data.csv");
        string[] csvLines =
        {
            "# This is a comment line and will be ignored by the parser",
            "Name;Age;Country",               // Header line (HasHeaders = true)
            "Alice;30;USA",
            "Bob;25;Canada",
            "Charlie;35;UK"
        };
        File.WriteAllLines(csvPath, csvLines);

        // 3. Configure CsvDataLoadOptions.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true) // true => first line has headers
        {
            Delimiter = ';',   // Custom column separator
            CommentChar = '#', // Lines starting with this character are treated as comments
            QuoteChar = '"'    // Default quote character (optional to set explicitly)
        };

        // 4. Create a CsvDataSource using the file path and the configured options.
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // 5. Load the template document (already saved) and build the report.
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the data source and the root name "persons".
        engine.BuildReport(doc, dataSource, "persons");

        // 6. Save the generated report.
        string resultPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(resultPath);

        // The example finishes without waiting for user input.
    }
}
