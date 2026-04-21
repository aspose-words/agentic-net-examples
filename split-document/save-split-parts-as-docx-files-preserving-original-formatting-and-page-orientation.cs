using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Prepare output folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputDir = Path.Combine(artifactsDir, "SplitParts");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑section document with different orientations.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1 – portrait.
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.Writeln("This is the content of Section 1 (portrait).");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 – landscape.
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.Writeln("This is the content of Section 2 (landscape).");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3 – portrait again.
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.Writeln("This is the content of Section 3 (portrait).");

        // Save the original document (optional, just for reference).
        string sourcePath = Path.Combine(artifactsDir, "SourceDocument.docx");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Split the document by sections, preserving formatting and orientation.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section srcSection = sourceDoc.Sections[i];

            // Create a new empty document that will hold the single section.
            Document partDoc = new Document();
            // Remove the automatically created empty section.
            partDoc.RemoveAllChildren();

            // Import the source section into the new document.
            // NodeImporter works with whole documents, not individual nodes.
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(srcSection, true);
            partDoc.AppendChild(importedSection);

            // Save the split part as DOCX.
            string partPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");
            partDoc.Save(partPath, SaveFormat.Docx);
        }

        // -----------------------------------------------------------------
        // 3. Validate that all split files were created.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string expectedPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");
            if (!File.Exists(expectedPath))
                throw new FileNotFoundException($"Expected split file not found: {expectedPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Document split into sections successfully. Files are located in:");
        Console.WriteLine(outputDir);
    }
}
