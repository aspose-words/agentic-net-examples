using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample document with mixed orientations.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First page – portrait (default).
        builder.Writeln("This is the portrait page.");
        builder.InsertBreak(BreakType.PageBreak);

        // Insert a new section so that the next page can have a different orientation.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("This is the landscape page.");

        // Save the source document.
        const string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Reload the document to work with it.
        // -----------------------------------------------------------------
        Document doc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // Split the document by sections and preserve each section's orientation.
        // -----------------------------------------------------------------
        int sectionIndex = 0;
        foreach (Section sec in doc.Sections)
        {
            // Create a new empty document and remove its default empty section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // ensures the imported section becomes the first one.

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(doc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sec, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string outPath = $"Section_{sectionIndex + 1}.docx";
            splitDoc.Save(outPath);

            // Validate that the file was created.
            if (!File.Exists(outPath))
                throw new Exception($"Failed to create split file for section {sectionIndex + 1}.");

            // Validate that the orientation matches the original.
            Document checkDoc = new Document(outPath);
            Orientation expected = sec.PageSetup.Orientation;
            Orientation actual = checkDoc.FirstSection.PageSetup.Orientation;
            if (expected != actual)
                throw new Exception($"Orientation mismatch on section {sectionIndex + 1}. Expected: {expected}, Actual: {actual}.");

            sectionIndex++;
        }

        // All sections split successfully.
        Console.WriteLine("Document split completed successfully.");
    }
}
