using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample EPUB with two chapters
        string epubPath = Path.Combine(outputDir, "sample.epub");
        CreateSampleEpub(epubPath);

        // Load the EPUB document
        Document srcDoc = new Document(epubPath);

        // Find all Heading 1 paragraphs (treated as chapter titles)
        List<Paragraph> chapterHeadings = new List<Paragraph>();
        foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                chapterHeadings.Add(para);
        }

        if (chapterHeadings.Count == 0)
            throw new InvalidOperationException("No chapter headings (Heading 1) were found in the EPUB.");

        // Split the document by chapters and save each as HTML
        for (int i = 0; i < chapterHeadings.Count; i++)
        {
            Paragraph startHeading = chapterHeadings[i];
            Paragraph endHeading = (i + 1 < chapterHeadings.Count) ? chapterHeadings[i + 1] : null;

            // Create a new empty document for the chapter
            Document chapterDoc = new Document();
            // Ensure the document has a section to receive imported nodes
            Section targetSection = chapterDoc.FirstSection;
            Body targetBody = targetSection.Body;

            // Prepare an importer to copy nodes from source to target
            NodeImporter importer = new NodeImporter(srcDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);

            // Iterate over nodes from startHeading up to (but not including) endHeading
            Node curNode = startHeading;
            while (curNode != null && curNode != endHeading)
            {
                Node importedNode = importer.ImportNode(curNode, true);
                targetBody.AppendChild(importedNode);
                curNode = curNode.NextSibling;
            }

            // Save the chapter as HTML
            string chapterFileName = Path.Combine(outputDir, $"Chapter_{i + 1}.html");
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Preserve layout as much as possible
            saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection;
            chapterDoc.Save(chapterFileName, saveOptions);

            // Validate that the file was created
            if (!File.Exists(chapterFileName))
                throw new InvalidOperationException($"Failed to create chapter file: {chapterFileName}");
        }
    }

    // Helper method to create a simple EPUB with two chapters
    private static void CreateSampleEpub(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of the first chapter. It contains some sample text.");

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the content of the second chapter. More sample text goes here.");

        // Save as EPUB
        doc.Save(path, SaveFormat.Epub);
    }
}
