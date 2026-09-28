using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with a portrait section and a landscape section.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First (portrait) section.
        builder.Writeln("This is the portrait section.");
        // Insert a section break to start a new section on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second (landscape) section.
        builder.CurrentSection.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("This is the landscape section.");

        // Save the source document.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // Split the document by sections, preserving each section's orientation.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section section = sourceDoc.Sections[i];

            // Create a new empty document to hold the imported section.
            Document partDoc = new Document();

            // Import the section from the source document into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);

            // Append the imported section to the new document.
            partDoc.AppendChild(importedSection);

            // Save the split part.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            partDoc.Save(partPath);

            // Validate that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split document: {partPath}");
        }

        // Validate that the expected number of split files exist.
        int expectedParts = sourceDoc.Sections.Count;
        for (int i = 1; i <= expectedParts; i++)
        {
            string partPath = Path.Combine(outputDir, $"Part_{i}.docx");
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Expected split file not found: {partPath}");
        }

        // All done.
        Console.WriteLine("Document split completed successfully.");
    }
}
