using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        const string sourcePath = "Source.docx";

        // Ensure any previous files are removed.
        if (File.Exists(sourcePath))
            File.Delete(sourcePath);
        for (int i = 1; i <= 10; i++)
        {
            string partPath = $"Section_{i}.docx";
            if (File.Exists(partPath))
                File.Delete(partPath);
        }

        // Create a sample document with two sections.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");
        sampleDoc.Save(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Split the document by sections.
        int sectionIndex = 1;
        foreach (Section section in sourceDoc.Sections)
        {
            // Create a new empty document.
            Document splitDoc = new Document();
            // Remove the default empty section that a new Document contains.
            splitDoc.RemoveAllChildren();

            // Import the section from the source document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = $"Section_{sectionIndex}.docx";
            splitDoc.Save(splitPath);

            // Validate that the file was created.
            if (!File.Exists(splitPath))
                throw new InvalidOperationException($"Failed to create split file: {splitPath}");

            sectionIndex++;
        }

        // Final validation: ensure at least one split file exists.
        if (sectionIndex == 1)
            throw new InvalidOperationException("No sections were found to split.");

        // Confirmation (non-interactive).
        Console.WriteLine($"Document split into {sectionIndex - 1} section file(s).");
    }
}
