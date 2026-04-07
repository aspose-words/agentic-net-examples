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

        // Paths for temporary files.
        const string csvPath = "people.csv";
        const string templatePath = "template.docx";
        const string outputPath = "ReportOutput.docx";

        // Create a sample CSV file with comment lines (starting with '#'), a header, and data rows.
        string[] csvLines =
        {
            "# This is a comment line that should be ignored",
            "# Another comment line",
            "Name,Age",
            "Alice,30",
            "Bob,25",
            "Charlie,35"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // Build a simple Word template containing LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template back from disk (demonstrates the load step).
        Document loadedTemplate = new Document(templatePath);

        // Configure CSV loading options to treat the first line as headers and ignore comment lines.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.CommentChar = '#'; // Lines starting with '#' will be skipped.
        loadOptions.Delimiter = ',';   // Use comma as the column separator.

        // Create a CsvDataSource from the CSV file stream using the configured options.
        using (FileStream csvStream = File.OpenRead(csvPath))
        {
            CsvDataSource csvDataSource = new CsvDataSource(csvStream, loadOptions);

            // Build the report by merging the CSV data into the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, csvDataSource, "persons");
        }

        // Save the generated report.
        loadedTemplate.Save(outputPath);
    }
}
