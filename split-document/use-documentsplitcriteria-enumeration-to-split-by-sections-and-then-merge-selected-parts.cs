using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitAndMergeBySection
{
    public static void Main()
    {
        // Paths for the documents.
        const string sourcePath = "Source.docx";
        const string mergedPath = "Merged.docx";

        // -----------------------------------------------------------------
        // 1. Create a sample document with three sections.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is the content of Section {i}.");
            // Insert a section break after each section except the last one.
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and split it by sections manually.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        List<Document> splitParts = new List<Document>();

        foreach (Section sec in loadedDoc.Sections)
        {
            // Create a new empty document for the current section.
            Document part = new Document();
            part.RemoveAllChildren(); // Ensure the document is empty.

            // Import the section from the source document into the new document.
            NodeImporter importer = new NodeImporter(loadedDoc, part, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sec, true);
            part.AppendChild(importedSection);

            splitParts.Add(part);
        }

        // Validate that we have the expected number of sections.
        if (splitParts.Count != 3)
            throw new InvalidOperationException($"Expected 3 split parts, but got {splitParts.Count}.");

        // -----------------------------------------------------------------
        // 3. Merge selected parts (e.g., first and third sections) into a new document.
        // -----------------------------------------------------------------
        Document mergedDoc = new Document();
        mergedDoc.RemoveAllChildren(); // Remove the default empty section.

        // Indices of parts to merge.
        int[] indicesToMerge = { 0, 2 }; // first and third sections.

        foreach (int index in indicesToMerge)
        {
            Document part = splitParts[index];
            // Import each section from the part into the merged document.
            NodeImporter importer = new NodeImporter(part, mergedDoc, ImportFormatMode.KeepSourceFormatting);
            foreach (Section sec in part.Sections)
            {
                Section importedSection = (Section)importer.ImportNode(sec, true);
                mergedDoc.AppendChild(importedSection);
            }
        }

        // Save the merged document.
        mergedDoc.Save(mergedPath);

        // -----------------------------------------------------------------
        // 4. Validate that the output files exist.
        // -----------------------------------------------------------------
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException($"Source document not found: {sourcePath}");
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException($"Merged document not found: {mergedPath}");

        Console.WriteLine("Document split and merge completed successfully.");
    }
}
