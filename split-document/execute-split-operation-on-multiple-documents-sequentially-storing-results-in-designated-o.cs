using System;
using System.IO;
using Aspose.Words;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define input and output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure clean environment.
        if (Directory.Exists(inputDir))
            Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample source documents.
        for (int docIndex = 1; docIndex <= 2; docIndex++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add three sections with distinct content.
            for (int sec = 1; sec <= 3; sec++)
            {
                // Write some body text.
                builder.Writeln($"Document {docIndex} - Section {sec} body content.");

                // Add a header to demonstrate preservation.
                builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
                builder.Writeln($"Header for Doc{docIndex} Sec{sec}");
                builder.MoveToDocumentEnd();

                // Insert a section break after each section except the last.
                if (sec < 3)
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
            }

            string sourcePath = Path.Combine(inputDir, $"SampleDoc{docIndex}.docx");
            doc.Save(sourcePath);
        }

        // Process each document: split by sections.
        foreach (string sourceFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document sourceDoc = new Document(sourceFile);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourceFile);

            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Prepare a new empty document.
                Document splitDoc = new Document();
                splitDoc.RemoveAllChildren(); // Remove the default empty section.

                // Import the current section from the source document.
                NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
                Section importedSection = (Section)importer.ImportNode(sourceDoc.Sections[i], true);
                splitDoc.AppendChild(importedSection);

                // Save the split document.
                string splitFileName = $"{fileNameWithoutExt}_Section{i + 1}.docx";
                string splitPath = Path.Combine(outputDir, splitFileName);
                splitDoc.Save(splitPath);
            }
        }

        // Validation: ensure each expected split file exists.
        foreach (string sourceFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourceFile);
            Document src = new Document(sourceFile);
            for (int i = 0; i < src.Sections.Count; i++)
            {
                string expectedPath = Path.Combine(outputDir, $"{fileNameWithoutExt}_Section{i + 1}.docx");
                if (!File.Exists(expectedPath))
                    throw new FileNotFoundException($"Expected split file not found: {expectedPath}");
            }
        }

        Console.WriteLine("Document splitting completed successfully.");
    }
}
