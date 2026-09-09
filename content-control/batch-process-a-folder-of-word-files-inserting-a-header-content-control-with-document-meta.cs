using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.BuildingBlocks;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output documents.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputDocs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample DOCX files if the input folder is empty.
        CreateSampleDocumentsIfNeeded(inputFolder);

        // Prepare a list to hold processing results for optional JSON reporting.
        var report = new List<object>();

        // Process each DOCX file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the document.
            var doc = new Document(filePath);

            // Retrieve some built‑in metadata.
            string title = doc.BuiltInDocumentProperties.Title ?? Path.GetFileNameWithoutExtension(filePath);
            string author = doc.BuiltInDocumentProperties.Author ?? "Unknown Author";

            // Build the metadata string that will be placed inside the header content control.
            string metadataText = $"Title: {title} | Author: {author}";

            // Move the builder to the primary header of the first section.
            var builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Insert an inline plain‑text content control.
            StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
            sdt.Title = "DocumentMetadata";
            sdt.Tag = "doc-metadata";

            // Clear any default children and add the metadata text.
            sdt.RemoveAllChildren();
            sdt.AppendChild(new Run(doc, metadataText));

            // Save the modified document to the output folder, preserving the original file name.
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            doc.Save(outputPath);

            // Record the result for the JSON summary.
            report.Add(new
            {
                FileName = Path.GetFileName(filePath),
                OutputPath = outputPath,
                Title = title,
                Author = author
            });
        }

        // Write a JSON summary of the batch operation.
        string jsonReportPath = Path.Combine(outputFolder, "summary.json");
        File.WriteAllText(jsonReportPath, JsonConvert.SerializeObject(report, Formatting.Indented));
    }

    private static void CreateSampleDocumentsIfNeeded(string folder)
    {
        // If the folder already contains DOCX files, assume samples are present.
        if (Directory.GetFiles(folder, "*.docx").Length > 0)
            return;

        for (int i = 1; i <= 2; i++)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln($"This is the content of sample document {i}.");
            doc.BuiltInDocumentProperties.Title = $"Sample Document {i}";
            doc.BuiltInDocumentProperties.Author = $"Author {i}";
            string filePath = Path.Combine(folder, $"Sample{i}.docx");
            doc.Save(filePath);
        }
    }
}
