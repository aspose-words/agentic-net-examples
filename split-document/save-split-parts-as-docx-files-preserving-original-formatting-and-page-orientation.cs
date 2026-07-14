using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with two sections having different page orientations.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = CreateSampleDocument();
        sourceDoc.Save(sourcePath);

        // Load the document.
        Document doc = new Document(sourcePath);

        // Split the document by sections, preserving formatting and orientation.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            // Create a new empty document for the current part.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(doc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(doc.Sections[i], true);
            partDoc.AppendChild(importedSection);

            // Save the split part as a DOCX file.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            partDoc.Save(partPath);

            // Verify that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split part: {partPath}");
        }
    }

    // Creates a sample document containing two sections with distinct page setups.
    private static Document CreateSampleDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section – default portrait orientation.
        builder.Writeln("This is the first section (Portrait).");
        builder.Writeln("It contains some sample text.");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Change page orientation for the second section to landscape.
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("This is the second section (Landscape).");
        builder.Writeln("Page orientation is preserved when splitting.");

        return doc;
    }
}
