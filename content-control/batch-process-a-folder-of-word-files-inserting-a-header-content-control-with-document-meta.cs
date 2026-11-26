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
        // Define input and output folders relative to the current working directory.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a few sample DOCX files to demonstrate batch processing.
        CreateSampleDocuments(inputDir);

        // Process each DOCX file: insert a header content control with document metadata.
        var report = new List<object>();

        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Ensure the primary header exists; create it if missing.
            HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header == null)
            {
                header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                doc.FirstSection.HeadersFooters.Add(header);
            }

            // Create a block-level rich text content control for metadata.
            StructuredDocumentTag metadataSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "DocMetadata",
                Tag = "metadata"
            };

            // Build a paragraph that will hold the metadata text.
            Paragraph metaParagraph = new Paragraph(doc);
            // Gather metadata values, handling possible nulls.
            string title = doc.BuiltInDocumentProperties.Title ?? string.Empty;
            string author = doc.BuiltInDocumentProperties.Author ?? string.Empty;
            string created = doc.BuiltInDocumentProperties.CreatedTime != DateTime.MinValue
                ? doc.BuiltInDocumentProperties.CreatedTime.ToString("u")
                : string.Empty;

            // Append runs with the metadata.
            metaParagraph.AppendChild(new Run(doc, $"Title: {title}"));
            metaParagraph.AppendChild(new Run(doc, "\n"));
            metaParagraph.AppendChild(new Run(doc, $"Author: {author}"));
            metaParagraph.AppendChild(new Run(doc, "\n"));
            metaParagraph.AppendChild(new Run(doc, $"Created: {created}"));

            // Add the paragraph to the content control.
            metadataSdt.AppendChild(metaParagraph);

            // Insert the content control at the beginning of the header.
            // If the header already has content, prepend; otherwise, just add.
            if (header.FirstChild != null)
            {
                header.InsertBefore(metadataSdt, header.FirstChild);
            }
            else
            {
                header.AppendChild(metadataSdt);
            }

            // Save the modified document to the output folder.
            string outputPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Record information for the JSON report.
            report.Add(new
            {
                FileName = Path.GetFileName(filePath),
                OutputPath = outputPath,
                Metadata = new
                {
                    Title = title,
                    Author = author,
                    Created = created
                }
            });
        }

        // Serialize the processing report to JSON.
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText(Path.Combine(outputDir, "processing_report.json"), jsonReport);
    }

    // Helper method to create a few sample documents.
    private static void CreateSampleDocuments(string folderPath)
    {
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            Paragraph para = new Paragraph(doc);
            para.AppendChild(new Run(doc, $"This is sample document {i}."));
            doc.FirstSection.Body.AppendChild(para);

            // Set some built‑in properties to demonstrate metadata insertion.
            doc.BuiltInDocumentProperties.Title = $"Sample Document {i}";
            doc.BuiltInDocumentProperties.Author = $"Author {i}";
            doc.BuiltInDocumentProperties.CreatedTime = DateTime.UtcNow.AddDays(-i);

            string filePath = Path.Combine(folderPath, $"Sample{i}.docx");
            doc.Save(filePath);
        }
    }
}
