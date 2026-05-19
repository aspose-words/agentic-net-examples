using System;
using System.IO;
using Aspose.Words;

public class MergeSplitDocuments
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with three sections.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Content of Section 1.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3.");

        // Save the source document (optional, just for reference).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the source document into separate files – one per section.
        // -----------------------------------------------------------------
        string[] partPaths = new string[sourceDoc.Sections.Count];
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // ensure it has no default section.

            // Import the i‑th section from the source document into the new document.
            // ImportNode clones the node and re‑parents it to the destination document.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true);
            partDoc.Sections.Add(importedSection);

            // Save the part.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            partDoc.Save(partPath);
            partPaths[i] = partPath;
        }

        // -----------------------------------------------------------------
        // 3. Load the split documents and merge them into a single document.
        // -----------------------------------------------------------------
        Document mergedDoc = new Document(); // starts with an empty section.

        foreach (string partPath in partPaths)
        {
            Document part = new Document(partPath);
            // Append each part while preserving its original formatting.
            mergedDoc.AppendDocument(part, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged result.
        string mergedPath = Path.Combine(outputDir, "Merged.docx");
        mergedDoc.Save(mergedPath);

        // -----------------------------------------------------------------
        // 4. Simple validation – ensure the merged file exists.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not created.", mergedPath);
    }
}
