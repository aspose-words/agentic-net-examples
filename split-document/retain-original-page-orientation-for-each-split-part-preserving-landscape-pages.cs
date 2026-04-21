using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class SplitDocumentPreserveOrientation
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with mixed portrait and landscape sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section – portrait (default).
        builder.Writeln("This is the first section (Portrait).");

        // Insert a new section and set its orientation to landscape.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("This is the second section (Landscape).");

        // Insert another section and revert to portrait.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.Writeln("This is the third section (Portrait).");

        // Save the source document (optional, for inspection).
        string sourcePath = Path.Combine(outputDir, "SourceDocument.docx");
        sourceDoc.Save(sourcePath);

        // Split the document by sections, preserving each section's page setup (orientation).
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the current section into the new document.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true);
            partDoc.AppendChild(importedSection);

            // Save the split part.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            partDoc.Save(partPath);
        }

        // Validation: ensure that each expected part file exists.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split part not found: {partPath}");
        }

        // Indicate successful completion (no console interaction required).
        // The program will exit automatically.
    }
}
