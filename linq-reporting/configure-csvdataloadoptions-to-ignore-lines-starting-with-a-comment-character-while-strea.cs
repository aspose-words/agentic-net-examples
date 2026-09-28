using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV file with comment lines (starting with '#') and data rows.
        const string csvPath = "sample.csv";
        File.WriteAllText(csvPath,
            "# This is a comment line and should be ignored\r\n" +
            "# Another comment\r\n" +
            "Name,Age\r\n" +
            "Alice,30\r\n" +
            "Bob,25\r\n" +
            "# End of data comment\r\n");

        // Create a Word template programmatically.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags to iterate over CSV rows.
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Configure CSV loading options to ignore comment lines starting with '#'.
        var csvOptions = new CsvDataLoadOptions
        {
            CommentChar = '#',
            HasHeaders = true
        };

        // Create CSV data source with the configured options.
        var csvDataSource = new CsvDataSource(csvPath, csvOptions);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, csvDataSource, "Data");

        // Save the generated report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
