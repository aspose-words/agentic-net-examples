using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath, "Column1,Column2\nJohn,\n,Smith\nAlice,Johnson\n");

        // Load CSV data into a list of Row objects.
        List<Row> rows = new();
        foreach (var line in File.ReadAllLines(csvPath))
        {
            // Skip header.
            if (line.StartsWith("Column1")) continue;
            var parts = line.Split(',');
            rows.Add(new Row
            {
                Column1 = parts.Length > 0 ? parts[0] : string.Empty,
                Column2 = parts.Length > 1 ? parts[1] : string.Empty
            });
        }

        // Wrap data for LINQ Reporting.
        ReportModel model = new() { Rows = rows };

        // Create the template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("<<foreach [row in Rows]>>");
        // Each row will be placed in its own paragraph.
        builder.Writeln("<<[row.Column1]>> <<[row.Column2]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document doc = new(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Remove empty paragraphs that may have been generated.
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (int i = paragraphs.Count - 1; i >= 0; i--)
        {
            Paragraph para = (Paragraph)paragraphs[i];
            // Trim whitespace and line breaks.
            string text = para.GetText().Trim();
            if (string.IsNullOrEmpty(text))
            {
                para.Remove();
            }
        }

        // Save the final document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Root model class for the report.
public class ReportModel
{
    public List<Row> Rows { get; set; } = new();
}

// Data row class representing each CSV record.
public class Row
{
    public string Column1 { get; set; } = string.Empty;
    public string Column2 { get; set; } = string.Empty;
}
