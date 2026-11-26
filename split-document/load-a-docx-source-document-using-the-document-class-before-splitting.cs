using System;
using System.IO;
using Aspose.Words;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Create a sample source document with two sections.
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

        // Save the source document to disk.
        const string sourceFileName = "Source.docx";
        sourceDoc.Save(sourceFileName);

        // Load the DOCX source document using the Document class.
        Document loadedDoc = new Document(sourceFileName);

        // Split the loaded document by its sections.
        int sectionNumber = 0;
        foreach (Section section in loadedDoc.Sections)
        {
            sectionNumber++;

            // Create a new empty document for the split part.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(loadedDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitFileName = $"Section_{sectionNumber}.docx";
            splitDoc.Save(splitFileName);

            // Validate that the split file was created.
            if (!File.Exists(splitFileName))
                throw new Exception($"Failed to create split file: {splitFileName}");
        }

        // Confirm that the expected number of split files exist.
        int expectedFiles = loadedDoc.Sections.Count;
        int actualFiles = 0;
        for (int i = 1; i <= expectedFiles; i++)
        {
            if (File.Exists($"Section_{i}.docx"))
                actualFiles++;
        }

        if (actualFiles != expectedFiles)
            throw new Exception("One or more split documents were not created as expected.");
    }
}
