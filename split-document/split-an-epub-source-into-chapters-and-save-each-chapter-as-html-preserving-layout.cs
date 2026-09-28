using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample EPUB with three chapters.
        string sourceEpubPath = "sample.epub";
        CreateSampleEpub(sourceEpubPath);

        // Load the EPUB document.
        Document epubDoc = new Document(sourceEpubPath);

        // Find all Heading 1 paragraphs (chapters).
        List<Paragraph> chapterHeadings = new List<Paragraph>();
        foreach (Paragraph para in epubDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                chapterHeadings.Add(para);
        }

        if (chapterHeadings.Count == 0)
            throw new InvalidOperationException("No chapter headings found in the EPUB document.");

        int chapterNumber = 1;
        foreach (Paragraph heading in chapterHeadings)
        {
            // Create a new empty document for the chapter.
            Document chapterDoc = new Document();

            // Use the first (and only) section that exists in a newly created document.
            Section chapterSection = chapterDoc.FirstSection;

            // Prepare an importer once for this chapter.
            NodeImporter importer = new NodeImporter(epubDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);

            // Import nodes from the heading up to (but not including) the next heading.
            Node currentNode = heading;
            while (currentNode != null)
            {
                // Stop before the next heading (except for the first heading itself).
                if (currentNode != heading &&
                    currentNode.NodeType == NodeType.Paragraph &&
                    ((Paragraph)currentNode).ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                {
                    break;
                }

                // Import the node into the chapter document.
                Node importedNode = importer.ImportNode(currentNode, true);
                chapterSection.Body.AppendChild(importedNode);

                currentNode = currentNode.NextSibling;
            }

            // Save the chapter as HTML.
            string htmlFileName = $"Chapter_{chapterNumber}.html";
            chapterDoc.Save(htmlFileName, SaveFormat.Html);
            chapterNumber++;
        }

        // Validate that the expected HTML files were created.
        for (int i = 1; i <= chapterHeadings.Count; i++)
        {
            string filePath = $"Chapter_{i}.html";
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Expected output file not found: {filePath}");
        }

        // Optional cleanup.
        // File.Delete(sourceEpubPath);
    }

    // Helper to create a simple EPUB with three chapters.
    private static void CreateSampleEpub(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            // Chapter heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {i}");

            // Chapter body.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of chapter {i}. It contains some sample text to demonstrate splitting.");
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        }

        // Save as EPUB.
        doc.Save(path, SaveFormat.Epub);
    }
}
