using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with two sections having different orientations.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        CreateSampleDocument(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Split the document by sections, preserving formatting and orientation.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document and remove the default empty section.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren();

            // Import the current section from the source document.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true, ImportFormatMode.KeepSourceFormatting);

            // Append the imported section to the new document.
            partDoc.AppendChild(importedSection);

            // Save the split part as DOCX.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            partDoc.Save(partPath, SaveFormat.Docx);

            // Validate that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split part: {partPath}");
        }

        // No console output required.
    }

    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section – portrait orientation.
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.Writeln("Section 1 – Portrait orientation.");
        builder.Writeln("This is some sample text in the first section.");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section – landscape orientation.
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("Section 2 – Landscape orientation.");
        builder.Writeln("This is some sample text in the second section.");

        // Save the sample document.
        doc.Save(filePath, SaveFormat.Docx);
    }
}
