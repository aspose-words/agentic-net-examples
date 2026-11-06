using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders in the current working directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample documents with built‑in metadata.
        for (int i = 1; i <= 3; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is the body of sample document {i}.");

            // Set built‑in properties for metadata.
            doc.BuiltInDocumentProperties.Title = $"Sample Document {i}";
            doc.BuiltInDocumentProperties.Author = $"Author {i}";

            string inputPath = Path.Combine(inputFolder, $"Sample{i}.docx");
            doc.Save(inputPath);
        }

        // Process each document in the input folder.
        string[] inputFiles = Directory.GetFiles(inputFolder, "*.docx");
        var summary = inputFiles.Select(filePath =>
        {
            Document doc = new Document(filePath);

            // Ensure the first section has a primary header.
            HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header == null)
            {
                header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                doc.FirstSection.HeadersFooters.Add(header);
            }

            // Create a block‑level rich‑text content control for metadata.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "DocumentMetadata",
                Tag = "doc-metadata"
            };
            sdt.RemoveAllChildren();

            // Build the metadata text.
            string metadataText = $"Title: {doc.BuiltInDocumentProperties.Title}, Author: {doc.BuiltInDocumentProperties.Author}";
            Paragraph paragraph = new Paragraph(doc);
            paragraph.AppendChild(new Run(doc, metadataText));
            sdt.AppendChild(paragraph);

            // Insert the content control at the beginning of the header.
            header.PrependChild(sdt);

            // Save the modified document.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Return information for the summary report.
            return new
            {
                FileName = Path.GetFileName(filePath),
                Metadata = metadataText,
                OutputPath = outputPath
            };
        }).ToList();

        // Write a JSON summary of the batch operation.
        string jsonReport = JsonConvert.SerializeObject(summary, Formatting.Indented);
        File.WriteAllText(Path.Combine(outputFolder, "BatchReport.json"), jsonReport);
    }
}
