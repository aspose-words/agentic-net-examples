using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string sourceFileName = "Source.docx";

        // Ensure any previous files are removed.
        if (File.Exists(sourceFileName))
            File.Delete(sourceFileName);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with two sections.
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

        // Save the source document to disk.
        sourceDoc.Save(sourceFileName);

        // -----------------------------------------------------------------
        // 2. Load the source document from disk.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourceFileName);

        // -----------------------------------------------------------------
        // 3. Define split criteria: split by each Section node.
        // -----------------------------------------------------------------
        int sectionIndex = 1;
        foreach (Section section in loadedDoc.Sections)
        {
            // Create a new empty document for the split part.
            Document splitDoc = new Document();

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(loadedDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);

            // Append the imported section to the split document.
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitFileName = $"Section_{sectionIndex}.docx";
            if (File.Exists(splitFileName))
                File.Delete(splitFileName);
            splitDoc.Save(splitFileName);

            // Validate that the file was created.
            if (!File.Exists(splitFileName))
                throw new InvalidOperationException($"Failed to create split file: {splitFileName}");

            sectionIndex++;
        }

        // -----------------------------------------------------------------
        // 4. Validation: ensure the expected number of split files exist.
        // -----------------------------------------------------------------
        int expectedSplitCount = loadedDoc.Sections.Count;
        for (int i = 1; i <= expectedSplitCount; i++)
        {
            string expectedFile = $"Section_{i}.docx";
            if (!File.Exists(expectedFile))
                throw new FileNotFoundException($"Expected split file not found: {expectedFile}");
        }

        // Program completed successfully.
    }
}
