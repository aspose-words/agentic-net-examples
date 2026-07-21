using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample documents if they do not exist
        CreateSampleDocuments(inputFolder);

        // Process each DOCX file
        var summary = new List<object>();
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Ensure a primary header exists
            HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header == null)
            {
                header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                doc.FirstSection.HeadersFooters.Add(header);
            }

            // Create a block‑level rich‑text content control
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "DocMetadata",
                Tag = "doc-metadata"
            };

            // Build metadata text
            string title = doc.BuiltInDocumentProperties.Title ?? "Untitled";
            string author = doc.BuiltInDocumentProperties.Author ?? "Unknown";
            string metadataText = $"Title: {title} | Author: {author}";

            // Add paragraph with metadata inside the content control
            Paragraph para = new Paragraph(doc);
            para.AppendChild(new Run(doc, metadataText));
            sdt.AppendChild(para);

            // Insert the content control into the header
            header.AppendChild(sdt);

            // Save processed document
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Record summary information
            summary.Add(new
            {
                FileName = Path.GetFileName(filePath),
                Title = title,
                Author = author,
                ProcessedUtc = DateTime.UtcNow
            });
        }

        // Write summary JSON
        string jsonPath = Path.Combine(outputFolder, "summary.json");
        string json = JsonConvert.SerializeObject(summary, Formatting.Indented);
        File.WriteAllText(jsonPath, json);
    }

    private static void CreateSampleDocuments(string folder)
    {
        for (int i = 1; i <= 3; i++)
        {
            string fileName = $"Sample{i}.docx";
            string filePath = Path.Combine(folder, fileName);
            if (File.Exists(filePath))
                continue;

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is the content of {fileName}.");
            // Set some built‑in properties for demonstration
            doc.BuiltInDocumentProperties.Title = $"Sample Document {i}";
            doc.BuiltInDocumentProperties.Author = $"Author {i}";
            doc.Save(filePath);
        }
    }
}
