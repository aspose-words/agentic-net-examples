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

        // Define file paths in the current directory.
        string templatePath = "Template.docx";
        string csvPath = "People.csv";
        string outputPath = "Report.docx";

        // Create a simple CSV file with a header and three rows.
        // The second row contains empty fields to demonstrate removal of empty paragraphs.
        File.WriteAllText(csvPath,
            "Name,Age\r\n" +
            "John Doe,30\r\n" +
            ",\r\n" +               // Empty row – will produce an empty paragraph.
            "Jane Smith,25\r\n",
            Encoding.UTF8);

        // Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a foreach tag that iterates over the CSV rows (exposed as 'persons').
        builder.Writeln("<<foreach [person in persons]>>");
        // Write the fields; each iteration creates its own paragraph.
        builder.Writeln("<<[person.Name]>> <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        template.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Configure CSV data source options (CSV has headers).
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        CsvDataSource dataSource = new CsvDataSource(csvPath, loadOptions);

        // Initialize the reporting engine and enable removal of empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report. The root object name must match the tag reference ('persons').
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
