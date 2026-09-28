using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document with two sections that have
        //    different page orientations (portrait and landscape).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section – default portrait orientation.
        builder.Writeln("First section – portrait orientation.");
        for (int i = 0; i < 30; i++)
            builder.Writeln($"Portrait line {i + 1}");

        // Insert a new section and set its orientation to landscape.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        sourceDoc.Sections[1].PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("Second section – landscape orientation.");
        for (int i = 0; i < 30; i++)
            builder.Writeln($"Landscape line {i + 1}");

        // Save the source document (optional, just for reference).
        string sourcePath = "SourceDocument.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the document by sections, preserving formatting and
        //    page orientation of each section.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section originalSection = sourceDoc.Sections[i];

            // Create a new empty document.
            Document splitDoc = new Document();
            // Remove the default empty section that a new Document contains.
            splitDoc.RemoveAllChildren();

            // Import the section from the source document into the new document,
            // keeping source formatting (including orientation, headers/footers, etc.).
            Section importedSection = (Section)splitDoc.ImportNode(
                originalSection, true, ImportFormatMode.KeepSourceFormatting);

            // Append the imported section as the sole section of the split document.
            splitDoc.AppendChild(importedSection);

            // Save each split part as a separate DOCX file.
            string outPath = $"Section_{i + 1}.docx";
            splitDoc.Save(outPath);
        }

        // -----------------------------------------------------------------
        // 3. Validate that each expected split file was created.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string outPath = $"Section_{i + 1}.docx";
            if (!File.Exists(outPath))
                throw new Exception($"Expected split file not found: {outPath}");
        }

        // -----------------------------------------------------------------
        // 4. Verify that the page orientation of each split document matches
        //    the original section's orientation.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string outPath = $"Section_{i + 1}.docx";
            Document splitDoc = new Document(outPath);

            Orientation originalOrientation = sourceDoc.Sections[i].PageSetup.Orientation;
            Orientation splitOrientation = splitDoc.Sections[0].PageSetup.Orientation;

            if (originalOrientation != splitOrientation)
                throw new Exception($"Orientation mismatch in {outPath}");
        }

        Console.WriteLine("Document split completed successfully.");
    }
}
