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
        // Define input and output folders relative to the current directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample documents if the input folder is empty.
        if (!Directory.EnumerateFiles(inputFolder, "*.docx").Any())
        {
            CreateSampleDocument(Path.Combine(inputFolder, "Sample1.docx"), "First sample document.");
            CreateSampleDocument(Path.Combine(inputFolder, "Sample2.docx"), "Second sample document.");
        }

        // Process each DOCX file in the input folder.
        var processedFiles = Directory.GetFiles(inputFolder, "*.docx")
            .Select(filePath =>
            {
                // Load the document.
                Document doc = new Document(filePath);

                // Ensure the document has at least one section.
                Section section = doc.FirstSection ?? new Section(doc);
                if (doc.FirstSection == null)
                {
                    doc.AppendChild(section);
                }

                // Get or create the primary header for the first section.
                HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }

                // Build a metadata string from built‑in document properties.
                string metadata = $"Title: {doc.BuiltInDocumentProperties.Title ?? "N/A"}; " +
                                  $"Author: {doc.BuiltInDocumentProperties.Author ?? "N/A"}; " +
                                  $"Created: {doc.BuiltInDocumentProperties.CreatedTime.ToString("yyyy-MM-dd")}";

                // Create a block‑level rich‑text content control for the header.
                StructuredDocumentTag headerSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
                {
                    Title = "DocMetadata",
                    Tag = "doc-metadata"
                };
                headerSdt.RemoveAllChildren();

                // Add a paragraph with the metadata text inside the content control.
                Paragraph para = new Paragraph(doc);
                para.AppendChild(new Run(doc, metadata));
                headerSdt.AppendChild(para);

                // Insert the content control at the beginning of the header.
                if (header.FirstChild != null)
                {
                    header.InsertBefore(headerSdt, header.FirstChild);
                }
                else
                {
                    header.AppendChild(headerSdt);
                }

                // Save the modified document to the output folder.
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                doc.Save(outputPath);

                // Return information for reporting.
                return new
                {
                    InputFile = Path.GetFileName(filePath),
                    OutputFile = Path.GetFileName(outputPath),
                    Metadata = metadata
                };
            })
            .ToList();

        // Write a JSON report of the processing results.
        string reportPath = Path.Combine(outputFolder, "ProcessingReport.json");
        File.WriteAllText(reportPath, JsonConvert.SerializeObject(processedFiles, Formatting.Indented));
    }

    // Helper method to create a simple DOCX file with a single paragraph.
    private static void CreateSampleDocument(string filePath, string paragraphText)
    {
        Document doc = new Document();
        Paragraph para = new Paragraph(doc);
        para.AppendChild(new Run(doc, paragraphText));
        doc.FirstSection.Body.AppendChild(para);
        doc.BuiltInDocumentProperties.Title = Path.GetFileNameWithoutExtension(filePath);
        doc.BuiltInDocumentProperties.Author = "AsposeSample";
        doc.Save(filePath);
    }
}
