using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV reading.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create sample CSV file (includes an empty line that should be ignored).
        string csvPath = Path.Combine(outputDir, "data.csv");
        File.WriteAllText(csvPath, "Name,Age\r\nAlice,30\r\nBob,25\r\n\r\nCharlie,35\r\n");

        // Build a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Heading.
        builder.Writeln("People Report");
        builder.Writeln();

        // Begin foreach over rows.
        builder.Writeln("<<foreach [row in rows]>>");
        // Write data fields.
        builder.Writeln("Name: <<[row.Name]>>");
        builder.Writeln("Age: <<[row.Age]>>");
        // Blank line (may become an empty paragraph after processing).
        builder.Writeln();
        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Configure CSV data source (first row contains headers).
        var csvOptions = new CsvDataLoadOptions { HasHeaders = true };
        CsvDataSource csvData = new CsvDataSource(csvPath, csvOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, csvData, "rows");

        // Remove any empty paragraphs that may have been generated.
        for (int i = doc.GetChildNodes(NodeType.Paragraph, true).Count - 1; i >= 0; i--)
        {
            Paragraph para = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[i];
            if (string.IsNullOrWhiteSpace(para.GetText()))
                para.Remove();
        }

        // Save the final document.
        string outputPath = Path.Combine(outputDir, "PeopleReport.docx");
        doc.Save(outputPath);
    }
}
