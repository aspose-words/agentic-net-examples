using System;
using System.IO;
using Aspose.Words;

public class MergeSplitDocuments
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the documents
        string sourcePath = Path.Combine(outputDir, "source.docx");
        string part1Path = Path.Combine(outputDir, "part1.docx");
        string part2Path = Path.Combine(outputDir, "part2.docx");
        string combinedPath = Path.Combine(outputDir, "combined.docx");

        // 1. Create a sample source document with two sections
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");

        sourceDoc.Save(sourcePath);

        // 2. Split the source document into separate documents by sections
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Document partDoc = new Document();
            // Remove the default empty section that a new Document contains
            partDoc.RemoveAllChildren();

            // Import the specific section from the source document
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedSection = importer.ImportNode(sourceDoc.Sections[i], true);
            partDoc.AppendChild(importedSection);

            // Save the split part
            string partPath = i == 0 ? part1Path : part2Path;
            partDoc.Save(partPath);
        }

        // 3. Load the split documents
        Document part1 = new Document(part1Path);
        Document part2 = new Document(part2Path);

        // 4. Merge the loaded documents using AppendDocument
        Document mergedDoc = new Document();
        mergedDoc.AppendDocument(part1, ImportFormatMode.KeepSourceFormatting);
        mergedDoc.AppendDocument(part2, ImportFormatMode.KeepSourceFormatting);

        // 5. Save the combined document
        mergedDoc.Save(combinedPath);

        // 6. Validate that the combined file exists
        if (!File.Exists(combinedPath))
        {
            throw new InvalidOperationException("Combined document was not created.");
        }
    }
}
