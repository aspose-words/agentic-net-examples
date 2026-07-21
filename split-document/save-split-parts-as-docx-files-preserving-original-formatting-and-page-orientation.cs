using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with two sections having different
        //    page orientations (portrait and landscape) and some text.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section – default portrait orientation.
        builder.Writeln("First section – portrait orientation.");
        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section – change orientation to landscape.
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("Second section – landscape orientation.");

        // Save the source document (optional, for inspection).
        string sourcePath = Path.Combine(outputDir, "SourceDocument.docx");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Split the source document by sections.
        //    Each section is saved as an independent DOCX file preserving
        //    its original formatting and page setup (including orientation).
        // -----------------------------------------------------------------
        int partNumber = 1;
        foreach (Section section in sourceDoc.Sections)
        {
            // Create a new empty document.
            Document partDoc = new Document();
            // Remove the automatically created empty section.
            partDoc.RemoveAllChildren();

            // Import the current section into the new document.
            // KeepSourceFormatting ensures that all formatting, headers,
            // footers, page setup, etc., are preserved.
            Node importedSection = partDoc.ImportNode(section, true, ImportFormatMode.KeepSourceFormatting);
            partDoc.AppendChild(importedSection);

            // Save the split part.
            string partPath = Path.Combine(outputDir, $"Part_{partNumber}.docx");
            partDoc.Save(partPath, SaveFormat.Docx);

            // Verify that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split part: {partPath}");

            partNumber++;
        }

        // -----------------------------------------------------------------
        // 3. Simple validation – ensure at least two parts were created.
        // -----------------------------------------------------------------
        int expectedParts = sourceDoc.Sections.Count;
        int actualParts = Directory.GetFiles(outputDir, "Part_*.docx").Length;
        if (actualParts != expectedParts)
            throw new InvalidOperationException($"Expected {expectedParts} split parts, but found {actualParts}.");

        // Execution completed without user interaction.
    }
}
