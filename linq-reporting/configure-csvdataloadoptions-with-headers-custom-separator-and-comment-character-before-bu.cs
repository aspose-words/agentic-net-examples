using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible non‑UTF8 CSV encoding.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a temporary folder for all generated files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create a sample CSV file with headers, a custom delimiter ';' and a comment line starting with '#'.
        string csvPath = Path.Combine(workDir, "People.csv");
        string[] csvLines =
        {
            "# This is a comment line and will be ignored by the parser",
            "Name;Age;Country",
            "Alice;30;USA",
            "Bob;25;Canada",
            "Charlie;35;UK"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // 2. Build a Word template that uses LINQ Reporting tags to iterate over the CSV data.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin a foreach loop over the data source named "persons".
        builder.Writeln("<<foreach [row in persons]>>");
        // Output each column value.
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Age: <<[row.Age]>>");
        builder.Writeln("Country: <<[row.Country]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Configure CSV loading options: headers present, ';' as delimiter, '#' as comment character.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true)
        {
            Delimiter = ';',
            CommentChar = '#',
            HasHeaders = true,
            QuoteChar = '"'
        };

        // 4. Create a CsvDataSource using the file and the configured options.
        CsvDataSource csvDataSource = new CsvDataSource(csvPath, loadOptions);

        // 5. Load the template document (demonstrating the load step).
        Document doc = new Document(templatePath);

        // 6. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name used in the template tags is "persons".
        engine.BuildReport(doc, csvDataSource, "persons");

        // 7. Save the generated report.
        string outputPath = Path.Combine(workDir, "PeopleReport.docx");
        doc.Save(outputPath);

        // The example finishes without waiting for user input.
    }
}
