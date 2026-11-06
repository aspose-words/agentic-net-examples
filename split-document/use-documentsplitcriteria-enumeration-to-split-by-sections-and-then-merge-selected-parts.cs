using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Ensure a clean environment.
        string[] filesToDelete = { "Source.docx", "Section_1.docx", "Section_2.docx", "Section_3.docx", "Merged.docx" };
        foreach (var f in filesToDelete)
        {
            if (File.Exists(f))
                File.Delete(f);
        }

        // 1. Create a sample document with three sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Section 1 - Header");
        builder.InsertParagraph();
        builder.Writeln("Content of the first section.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - Header");
        builder.InsertParagraph();
        builder.Writeln("Content of the second section.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - Header");
        builder.InsertParagraph();
        builder.Writeln("Content of the third section.");

        sourceDoc.Save("Source.docx");

        // 2. Split the document by sections.
        int sectionIndex = 1;
        foreach (Section sec in sourceDoc.Sections)
        {
            Document splitDoc = new Document();
            // Remove the default empty section that a new Document contains.
            splitDoc.RemoveAllChildren();

            // Import the section into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sec, true);
            splitDoc.AppendChild(importedSection);

            string splitFileName = $"Section_{sectionIndex}.docx";
            splitDoc.Save(splitFileName);
            sectionIndex++;
        }

        // Verify split files exist.
        for (int i = 1; i <= sourceDoc.Sections.Count; i++)
        {
            string path = $"Section_{i}.docx";
            if (!File.Exists(path))
                throw new Exception($"Expected split file not found: {path}");
        }

        // 3. Merge selected sections (e.g., Section 1 and Section 3) into a new document.
        Document mergedDoc = new Document();
        mergedDoc.RemoveAllChildren(); // Remove the default empty section.

        int[] sectionsToMerge = { 0, 2 }; // Zero‑based indices for Section 1 and Section 3.
        foreach (int idx in sectionsToMerge)
        {
            if (idx < 0 || idx >= sourceDoc.Sections.Count)
                throw new Exception($"Section index out of range: {idx}");

            Section sec = sourceDoc.Sections[idx];
            NodeImporter importer = new NodeImporter(sourceDoc, mergedDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSec = (Section)importer.ImportNode(sec, true);
            mergedDoc.AppendChild(importedSec);
        }

        mergedDoc.Save("Merged.docx");

        // Verify merged file exists.
        if (!File.Exists("Merged.docx"))
            throw new Exception("Merged document was not created.");

        // Simple confirmation.
        Console.WriteLine("Document splitting and merging completed successfully.");
    }
}
