using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputDir = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents if the input folder is empty.
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            CreateSampleDocument(Path.Combine(inputDir, "Sample1.docx"), "First sample document.");
            CreateSampleDocument(Path.Combine(inputDir, "Sample2.docx"), "Second sample document.");
        }

        // Process each DOCX file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);

            // Ensure word count properties are up‑to‑date.
            doc.UpdateWordCount();

            // Gather metadata from built‑in properties.
            var metadata = new
            {
                Title = doc.BuiltInDocumentProperties.Title,
                Author = doc.BuiltInDocumentProperties.Author,
                Created = doc.BuiltInDocumentProperties.CreatedTime,
                Words = doc.BuiltInDocumentProperties.Words,
                Pages = doc.BuiltInDocumentProperties.Pages
            };

            // Serialize metadata to a formatted JSON string.
            string json = JsonConvert.SerializeObject(metadata, Formatting.Indented);

            // Ensure the primary header exists and obtain a reference to it.
            HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header == null)
            {
                header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                doc.FirstSection.HeadersFooters.Add(header);
            }

            // Create a block‑level rich‑text content control.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "DocMetadata",
                Tag = "metadata"
            };

            // Add a paragraph with the JSON metadata inside the content control.
            Paragraph para = new Paragraph(doc);
            para.AppendChild(new Run(doc, json));
            sdt.AppendChild(para);

            // Insert the content control into the header.
            header.AppendChild(sdt);

            // Save the modified document to the output folder.
            string outPath = Path.Combine(outputDir, Path.GetFileName(filePath));
            doc.Save(outPath);
        }
    }

    // Helper method to create a simple document with some built‑in properties.
    private static void CreateSampleDocument(string path, string bodyText)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(bodyText);

        // Set sample built‑in properties.
        doc.BuiltInDocumentProperties.Title = Path.GetFileNameWithoutExtension(path);
        doc.BuiltInDocumentProperties.Author = "Aspose Sample";

        doc.Save(path);
    }
}
