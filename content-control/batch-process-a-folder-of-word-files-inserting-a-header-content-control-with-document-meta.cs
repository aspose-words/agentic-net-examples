using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputDocs");

        // Ensure folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample documents
        CreateSampleDocuments(inputFolder);

        // Process each .docx file
        var report = new List<DocumentReport>();
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(filePath);

            // Ensure header exists
            HeaderFooter header = GetOrCreateHeader(doc);

            // Insert header content control with metadata
            InsertHeaderMetadataControl(doc, header);

            // Save processed document
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Add entry to report
            report.Add(new DocumentReport
            {
                FileName = Path.GetFileName(filePath),
                Title = doc.BuiltInDocumentProperties.Title,
                Author = doc.BuiltInDocumentProperties.Author
            });
        }

        // Write JSON report
        string jsonReportPath = Path.Combine(outputFolder, "processing_report.json");
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText(jsonReportPath, json);
    }

    private static void CreateSampleDocuments(string folder)
    {
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            doc.BuiltInDocumentProperties.Title = $"Sample Document {i}";
            doc.BuiltInDocumentProperties.Author = $"Author {i}";

            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, $"This is the content of sample document {i}.");
            para.AppendChild(run);
            doc.FirstSection.Body.AppendChild(para);

            string filePath = Path.Combine(folder, $"Sample{i}.docx");
            doc.Save(filePath);
        }
    }

    private static HeaderFooter GetOrCreateHeader(Document doc)
    {
        Section firstSection = doc.FirstSection;
        HeaderFooterCollection headers = firstSection.HeadersFooters;
        HeaderFooter header = headers[HeaderFooterType.HeaderPrimary];
        if (header == null)
        {
            header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            headers.Add(header);
        }
        return header;
    }

    private static void InsertHeaderMetadataControl(Document doc, HeaderFooter header)
    {
        // Create block-level rich text content control
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "DocumentMetadata",
            Tag = "doc-metadata"
        };

        // Build metadata text
        string title = doc.BuiltInDocumentProperties.Title ?? "N/A";
        string author = doc.BuiltInDocumentProperties.Author ?? "N/A";
        string metadataText = $"Title: {title}; Author: {author}";

        // Add paragraph with metadata inside the content control
        Paragraph para = new Paragraph(doc);
        Run run = new Run(doc, metadataText);
        para.AppendChild(run);
        sdt.AppendChild(para);

        // Insert the content control at the beginning of the header
        Node firstNode = header.FirstChild;
        if (firstNode != null)
        {
            header.InsertBefore(sdt, firstNode);
        }
        else
        {
            header.AppendChild(sdt);
        }
    }

    private class DocumentReport
    {
        public string FileName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
    }
}
