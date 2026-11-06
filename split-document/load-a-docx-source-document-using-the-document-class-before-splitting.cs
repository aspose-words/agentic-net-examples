using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the sample source document.
        const string sourcePath = "Source.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX document with two sections.
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
        // Step 2: Load the DOCX source document using the Document class.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // Step 3: Split the loaded document by its sections.
        // -----------------------------------------------------------------
        int sectionIndex = 1;
        foreach (Section section in loadedDoc.Sections)
        {
            // Create a new empty document for the split part.
            Document splitDoc = new Document();

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(loadedDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            splitDoc.AppendChild(importedSection);

            // Define the output file name for this section.
            string splitPath = $"Section{sectionIndex}.docx";

            // Save the split document.
            splitDoc.Save(splitPath);

            // Validate that the file was created.
            if (!File.Exists(splitPath))
                throw new InvalidOperationException($"Failed to create split file: {splitPath}");

            sectionIndex++;
        }

        // -----------------------------------------------------------------
        // Step 4: Verify that the expected split files exist.
        // -----------------------------------------------------------------
        for (int i = 1; i < sectionIndex; i++)
        {
            string expectedPath = $"Section{i}.docx";
            if (!File.Exists(expectedPath))
                throw new FileNotFoundException($"Expected split document not found: {expectedPath}");
        }

        // All operations completed successfully.
    }
}
