using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Define input and output folders relative to the working directory.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample documents if the input folder is empty.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder sampleBuilder = new DocumentBuilder(sampleDoc);

                // Add some body text.
                sampleBuilder.Writeln($"This is sample document {i}.");

                // Set built‑in document properties that we will later insert into the header.
                sampleDoc.BuiltInDocumentProperties.Title = $"Sample Title {i}";
                sampleDoc.BuiltInDocumentProperties.Author = $"Author {i}";

                string samplePath = Path.Combine(inputFolder, $"Doc{i}.docx");
                sampleDoc.Save(samplePath);
            }
        }

        // Process each DOCX file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Iterate through all sections to add a header content control.
            for (int secIndex = 0; secIndex < doc.Sections.Count; secIndex++)
            {
                // Move the builder to the current section.
                builder.MoveToSection(secIndex);

                // Move to the primary header of the current section (creates it if missing).
                builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

                // Insert a plain‑text block‑level content control.
                StructuredDocumentTag headerSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
                headerSdt.Title = "DocumentMetadata";
                headerSdt.Tag = "DocMetadata";

                // Build the metadata text.
                string title = doc.BuiltInDocumentProperties.Title ?? string.Empty;
                string author = doc.BuiltInDocumentProperties.Author ?? string.Empty;
                string metadataText = $"Title: {title} | Author: {author}";

                // Write the metadata inside the content control.
                builder.Write(metadataText);
            }

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);
        }
    }
}
