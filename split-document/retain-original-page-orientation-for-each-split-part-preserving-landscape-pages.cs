using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a portrait page followed by a landscape page.
        string sourcePath = "Source.docx";
        CreateSampleDocument(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);
        int sectionCount = sourceDoc.Sections.Count;

        // Split each section (each page) into a separate document while preserving its orientation.
        for (int i = 0; i < sectionCount; i++)
        {
            Section srcSection = sourceDoc.Sections[i];

            // Create a new empty document and remove the default empty section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren();

            // Import the source section into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(srcSection, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string outPath = $"Page_{i + 1}.docx";
            splitDoc.Save(outPath);

            // Verify the file was created.
            if (!File.Exists(outPath))
                throw new InvalidOperationException($"Failed to save split document for page {i + 1}.");

            // Verify orientation is preserved.
            Orientation actualOrientation = splitDoc.Sections[0].PageSetup.Orientation;
            Orientation expectedOrientation = srcSection.PageSetup.Orientation;
            if (actualOrientation != expectedOrientation)
                throw new InvalidOperationException(
                    $"Orientation mismatch on page {i + 1}. Expected {expectedOrientation}, got {actualOrientation}.");
        }
    }

    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section – portrait orientation.
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.Writeln("This is a portrait page.");
        // Force a new section on a new page.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section – landscape orientation.
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("This is a landscape page.");

        doc.Save(path);
    }
}
