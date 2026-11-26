using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Reporting;

public class LinkInfo
{
    public string Text { get; set; } = "";
    public string Url { get; set; } = "";
}

public class ReportModel
{
    public List<LinkInfo> Links { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample files.
        string existingFile = "existing.txt";
        File.WriteAllText(existingFile, "This file exists.");

        // Prepare data model.
        var model = new ReportModel
        {
            Links = new List<LinkInfo>
            {
                new LinkInfo { Text = "Existing File", Url = Path.GetFullPath(existingFile) },
                new LinkInfo { Text = "Missing File", Url = Path.GetFullPath("missing.txt") }
            }
        };

        // Create template document with LINQ Reporting tags.
        var templatePath = "template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [link in Links]>>");
        builder.Writeln("<<link [link.Url] [link.Text]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load template and build report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");
        var outputPath = "output.docx";
        doc.Save(outputPath);

        // Update fields to ensure hyperlink results are current.
        doc.UpdateFields();

        // Scan for broken hyperlinks.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldHyperlink)
            {
                var hyperlink = (FieldHyperlink)field;
                string address = hyperlink.Address ?? string.Empty;

                // Simple check: treat non‑HTTP addresses as file paths.
                bool isHttp = address.StartsWith("http", StringComparison.OrdinalIgnoreCase);
                bool targetExists = isHttp || File.Exists(address);

                if (!targetExists)
                {
                    string displayText = hyperlink.Result ?? string.Empty;
                    Console.WriteLine($"Broken hyperlink detected: Text=\"{displayText}\", Address=\"{address}\"");
                }
            }
        }
    }
}
