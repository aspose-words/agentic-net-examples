using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // Create a sample document with heading paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add three chapters, each starting with a Heading 1 style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 3.");

        // Save the source document for inspection (optional).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // ----------- Split the document by heading paragraphs -------------
        // Manually create separate Document objects for each heading section.
        List<Document> splitDocs = new List<Document>();

        // Get all paragraphs in the source document (deep traversal).
        NodeCollection paragraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);

        Document currentPart = null;
        NodeImporter importer = null;

        foreach (Paragraph para in paragraphs)
        {
            // Determine whether the paragraph is a heading.
            bool isHeading = para.ParagraphFormat.IsHeading;

            if (isHeading)
            {
                // When a new heading is encountered, finish the previous part.
                if (currentPart != null)
                    splitDocs.Add(currentPart);

                // Start a new part document.
                currentPart = new Document();
                currentPart.EnsureMinimum(); // Guarantees a section/body/paragraph exists.
                importer = new NodeImporter(sourceDoc, currentPart, ImportFormatMode.KeepSourceFormatting);
            }

            // If we have an active part, import the current paragraph into it.
            if (currentPart != null && importer != null)
            {
                Node importedNode = importer.ImportNode(para, true);
                currentPart.FirstSection.Body.AppendChild(importedNode);
            }
        }

        // Add the final part if it exists.
        if (currentPart != null)
            splitDocs.Add(currentPart);

        // Save each split part to a separate file.
        for (int i = 0; i < splitDocs.Count; i++)
        {
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            splitDocs[i].Save(partPath);
        }

        // Simple validation: ensure that the expected number of parts were created.
        int expectedParts = 3; // We added three heading paragraphs.
        if (splitDocs.Count != expectedParts)
            throw new InvalidOperationException($"Expected {expectedParts} split parts, but got {splitDocs.Count}.");

        // Verify that each output file exists.
        for (int i = 0; i < expectedParts; i++)
        {
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            if (!File.Exists(partPath))
                throw new FileNotFoundException($"Split part file not found: {partPath}");
        }

        Console.WriteLine("Document split completed successfully.");
    }
}
