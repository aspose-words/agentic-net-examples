using System;
using System.IO;
using Aspose.Words;

public class MergeSplitDocuments
{
    public static void Main()
    {
        // Create a sample source document with three sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Content of Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Content of Section 3");

        // Save the source document (optional, for reference).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Split the source document into separate documents per section.
        string[] partPaths = new string[sourceDoc.Sections.Count];
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document partDoc = new Document();

            // Import the current section into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sourceDoc.Sections[i], true);
            partDoc.AppendChild(importedSection);

            // Save the split part.
            string partPath = $"Part{i + 1}.docx";
            partDoc.Save(partPath);
            partPaths[i] = partPath;
        }

        // Verify that split parts were created.
        foreach (string path in partPaths)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"Expected split part not found: {path}");
        }

        // Load selected split documents (e.g., Part1 and Part3) and merge them.
        string part1Path = partPaths[0];
        string part3Path = partPaths[2];

        Document mergedDoc = new Document(part1Path);
        Document part3Doc = new Document(part3Path);
        mergedDoc.AppendDocument(part3Doc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        string combinedPath = "Combined.docx";
        mergedDoc.Save(combinedPath);

        // Validate that the combined document exists.
        if (!File.Exists(combinedPath))
            throw new FileNotFoundException($"Combined document was not created: {combinedPath}");

        // Optional cleanup (commented out).
        // File.Delete(sourcePath);
        // foreach (string path in partPaths) File.Delete(path);
    }
}
