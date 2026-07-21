using System;
using System.Collections.Generic;
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

        // Prepare sample CSV file.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.csv");
        File.WriteAllText(csvPath, "Name,Comment\r\nAlice,Hello\r\nBob,\r\nCharlie,World");

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // First section – will be populated from CSV and may contain empty paragraphs.
        builder.Writeln("Section 1 – CSV data:");
        builder.Writeln("<<foreach [row in data]>>");
        // Paragraph that can become empty if Comment column is empty.
        builder.Writeln("<<[row.Comment]>>");
        builder.Writeln("<</foreach>>");

        // Start a new section that must stay unchanged.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 – static content that must remain.");
        builder.Writeln("This paragraph should stay even if empty.");

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Configure CSV data source with headers.
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true); // first line contains column names
        CsvDataSource csvData = new CsvDataSource(csvPath, loadOptions);

        // Build the report without automatic empty‑paragraph removal.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(report, csvData, "data");

        // Remove empty paragraphs only from the first section (the CSV‑processed part).
        Section firstSection = report.Sections[0];
        List<Paragraph> emptyParagraphs = new();
        foreach (Paragraph para in firstSection.Body.Paragraphs)
        {
            if (string.IsNullOrWhiteSpace(para.GetText()))
                emptyParagraphs.Add(para);
        }
        foreach (Paragraph para in emptyParagraphs)
            para.Remove();

        // Save the final document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(outputPath);
    }
}
