using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three sections having different page orientations.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1 – Portrait (default orientation).
        builder.Writeln("Section 1 - Portrait");

        // Section 2 – Landscape.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("Section 2 - Landscape");

        // Section 3 – Portrait again.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.Writeln("Section 3 - Portrait again");

        // Split the document by sections, preserving formatting and orientation.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section originalSection = sourceDoc.Sections[i];

            // Create a new empty document to hold the single section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the section from the source document into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(originalSection, true);
            splitDoc.AppendChild(importedSection);

            // Save the split part as a DOCX file.
            string fileName = $"Section_{i + 1}.docx";
            splitDoc.Save(fileName, SaveFormat.Docx);

            // Verify that the file was created.
            if (!File.Exists(fileName))
                throw new Exception($"Failed to create split file: {fileName}");
        }
    }
}
