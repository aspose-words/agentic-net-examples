using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string inputFolder = Path.Combine(Environment.CurrentDirectory, "InputDocs");
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "OutputDocs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample documents if the input folder is empty.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"), "First Document", "Alice", new DateTime(2023, 1, 15));
            CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"), "Second Document", "Bob", new DateTime(2023, 2, 20));
            CreateSampleDocument(Path.Combine(inputFolder, "Sample3.docx"), "Third Document", "Carol", new DateTime(2023, 3, 25));
        }

        // Collect processing results for optional JSON report.
        var report = new List<ProcessedFileInfo>();

        // Process each DOCX file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Ensure a primary header exists.
            Section firstSection = doc.FirstSection;
            HeaderFooter header = firstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header == null)
            {
                header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                firstSection.HeadersFooters.Add(header);
            }

            // Build a block‑level rich‑text content control in the header.
            StructuredDocumentTag metaSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "DocumentMetadata",
                Tag = "DocMeta"
            };

            // Title paragraph.
            Paragraph titlePara = new Paragraph(doc);
            titlePara.AppendChild(new Run(doc, $"Title: {doc.BuiltInDocumentProperties.Title}"));
            metaSdt.AppendChild(titlePara);

            // Author paragraph.
            Paragraph authorPara = new Paragraph(doc);
            authorPara.AppendChild(new Run(doc, $"Author: {doc.BuiltInDocumentProperties.Author}"));
            metaSdt.AppendChild(authorPara);

            // Created date paragraph (UTC).
            Paragraph createdPara = new Paragraph(doc);
            createdPara.AppendChild(new Run(doc, $"Created: {doc.BuiltInDocumentProperties.CreatedTime:u}"));
            metaSdt.AppendChild(createdPara);

            // Insert the content control into the header.
            header.AppendChild(metaSdt);

            // Save the modified document to the output folder.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Record information for the report.
            report.Add(new ProcessedFileInfo
            {
                FileName = Path.GetFileName(filePath),
                Title = doc.BuiltInDocumentProperties.Title,
                Author = doc.BuiltInDocumentProperties.Author,
                CreatedUtc = doc.BuiltInDocumentProperties.CreatedTime
            });
        }

        // Write a JSON summary of the processed files.
        string jsonReportPath = Path.Combine(outputFolder, "ProcessingReport.json");
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText(jsonReportPath, json);
    }

    // Helper to create a simple document with some built‑in properties.
    private static void CreateSampleDocument(string path, string title, string author, DateTime created)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln($"This is the content of \"{title}\".");
        doc.BuiltInDocumentProperties.Title = title;
        doc.BuiltInDocumentProperties.Author = author;
        doc.BuiltInDocumentProperties.CreatedTime = created;
        doc.Save(path);
    }

    // DTO for the JSON report.
    private class ProcessedFileInfo
    {
        public string FileName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public DateTime CreatedUtc { get; set; }
    }
}
