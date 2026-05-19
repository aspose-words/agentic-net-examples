using System;
using System.IO;
using Aspose.Words;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define paths for the source and output files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);
        string sourcePath = Path.Combine(dataDir, "Source.docx");
        string outputDir = Path.Combine(dataDir, "SplitParts");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document with multiple sections.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section content.
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section content.
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");

        // Save the sample source document.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX source document using the Document class.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Split the loaded document by its sections.
        // -----------------------------------------------------------------
        for (int i = 0; i < loadedDoc.Sections.Count; i++)
        {
            // Create a new empty document that will hold the split part.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // Ensure the document is empty.

            // Use NodeImporter for efficient node import.
            NodeImporter importer = new NodeImporter(loadedDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);

            // Import the current section into the new document.
            Section importedSection = (Section)importer.ImportNode(loadedDoc.Sections[i], true);
            splitDoc.AppendChild(importedSection);

            // Save the split part.
            string partPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");
            splitDoc.Save(partPath);

            // Simple validation: ensure the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split part: {partPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Document split completed successfully.");
    }
}
