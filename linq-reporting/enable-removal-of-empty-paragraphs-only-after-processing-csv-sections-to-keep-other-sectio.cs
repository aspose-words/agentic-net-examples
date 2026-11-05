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

        // Prepare sample CSV data.
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath, "Name,Value\r\nAlice,100\r\n,200\r\nBob,300\r\n", Encoding.UTF8);

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Static content before CSV section.
        builder.Writeln("=== Report Start ===");

        // CSV processing section using LINQ Reporting tags.
        builder.Writeln("<<foreach [row in csv]>>");
        // Paragraph that may become empty if both columns are empty.
        builder.Writeln("<<[row.Name]>> - <<[row.Value]>>");
        builder.Writeln("<</foreach>>");

        // Static content after CSV section.
        builder.Writeln("=== Report End ===");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Load CSV data source with headers.
        var csvLoadOptions = new CsvDataLoadOptions { HasHeaders = true };
        var csvDataSource = new CsvDataSource(csvPath, csvLoadOptions);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, csvDataSource, "csv");

        // Remove empty paragraphs that may have been generated from empty CSV rows.
        // This is performed after the CSV data has been processed, leaving other sections untouched.
        RemoveEmptyParagraphs(reportDoc);

        // Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }

    // Removes paragraphs that contain only whitespace or are completely empty.
    private static void RemoveEmptyParagraphs(Document doc)
    {
        foreach (Section section in doc.Sections)
        {
            Body body = section.Body;
            // Iterate backwards because removing paragraphs changes the collection.
            for (int i = body.Paragraphs.Count - 1; i >= 0; i--)
            {
                Paragraph para = body.Paragraphs[i];
                // GetText includes the paragraph break, so trim to check content.
                if (string.IsNullOrWhiteSpace(para.GetText()))
                {
                    para.Remove();
                }
            }
        }
    }
}
