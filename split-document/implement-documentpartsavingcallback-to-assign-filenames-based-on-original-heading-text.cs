using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // 1. Create a sample document with Heading 1 paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First chapter
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter One");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the first chapter.");

        // Second chapter
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter Two");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the second chapter.");

        // Third chapter
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter Three");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the third chapter.");

        // Save the source document (optional, just to have a file on disk).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // 2. Load the document (simulating a real scenario where the file already exists).
        Document doc = new Document(sourcePath);

        // 3. Find all Heading 1 paragraphs.
        List<Paragraph> headingParagraphs = doc.GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Where(p => p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            .ToList();

        if (!headingParagraphs.Any())
            throw new InvalidOperationException("No Heading 1 paragraphs found to split the document.");

        // 4. Iterate over each heading and create a separate document containing that heading and its following content.
        for (int i = 0; i < headingParagraphs.Count; i++)
        {
            Paragraph heading = headingParagraphs[i];
            string headingText = heading.GetText().Trim();

            // Determine the node range for this chapter.
            Node startNode = heading;
            Node endNode = (i + 1 < headingParagraphs.Count) ? headingParagraphs[i + 1] : null;

            // Create a new empty document.
            Document splitDoc = new Document();

            // Import nodes belonging to the current chapter.
            Node currentNode = startNode;
            while (currentNode != null && currentNode != endNode)
            {
                Node importedNode = splitDoc.ImportNode(currentNode, true);
                splitDoc.FirstSection.Body.AppendChild(importedNode);
                currentNode = currentNode.NextSibling;
            }

            // 5. Build a safe filename from the heading text.
            string safeFileName = string.Concat(headingText.Split(Path.GetInvalidFileNameChars()))
                                      .Replace(' ', '_');
            string outputPath = $"{safeFileName}.docx";

            // 6. Save the split document.
            splitDoc.Save(outputPath, SaveFormat.Docx);

            // 7. Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create split file: {outputPath}");
        }

        // All split documents have been created successfully.
    }
}
