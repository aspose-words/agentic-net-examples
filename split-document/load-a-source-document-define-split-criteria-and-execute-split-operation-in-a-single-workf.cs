using System;
using System.IO;
using Aspose.Words;

public class SplitDocumentExample
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with three sections.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // Save the source document locally.
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the source document.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -------------------------------------------------
        // 3. Split the document by its sections.
        // -------------------------------------------------
        for (int i = 0; i < loadedDoc.Sections.Count; i++)
        {
            Section sourceSection = loadedDoc.Sections[i];

            // Create a new empty document for the current section.
            Document splitDoc = new Document();
            splitDoc.Sections.Clear(); // Remove the default empty section.

            // Import the section from the source document into the new document.
            NodeImporter importer = new NodeImporter(loadedDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sourceSection, true);

            // Add the imported section to the split document.
            splitDoc.Sections.Add(importedSection);

            // Save the split document.
            string splitPath = $"Section_{i + 1}.docx";
            splitDoc.Save(splitPath);

            // Verify that the file was created.
            if (!File.Exists(splitPath))
                throw new Exception($"Failed to create split file: {splitPath}");
        }

        // -------------------------------------------------
        // 4. Indicate successful completion.
        // -------------------------------------------------
        Console.WriteLine("Document split completed successfully.");
    }
}
