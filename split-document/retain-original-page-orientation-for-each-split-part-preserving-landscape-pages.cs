using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with mixed page orientations.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section – default portrait orientation.
        builder.Writeln("This is the first (portrait) section.");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Change orientation to landscape for the second section.
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("This is the second (landscape) section.");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        Directory.CreateDirectory(outputDir);

        // Split the document by sections, preserving each section's orientation.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section srcSection = sourceDoc.Sections[i];

            // Create a new empty document for the split part.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the source section into the new document.
            // NodeImporter works with whole documents, so we pass the source document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(srcSection, true);
            splitDoc.AppendChild(importedSection);

            // Save the split part.
            string outPath = Path.Combine(outputDir, $"SplitPart_{i + 1}.docx");
            splitDoc.Save(outPath, SaveFormat.Docx);

            // Validate that the orientation was retained.
            Orientation expected = srcSection.PageSetup.Orientation;
            Orientation actual = splitDoc.Sections[0].PageSetup.Orientation;
            if (expected != actual)
                throw new InvalidOperationException($"Orientation mismatch in part {i + 1}.");
        }

        // Verify that the expected number of split files were created.
        string[] files = Directory.GetFiles(outputDir, "*.docx");
        if (files.Length != sourceDoc.Sections.Count)
            throw new InvalidOperationException("Unexpected number of split output files.");

        Console.WriteLine("Document split completed successfully. Output files are located at:");
        Console.WriteLine(outputDir);
    }
}
