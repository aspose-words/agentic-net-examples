using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Sample HTML with chapters (using <h1> as chapter headings)
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8' />
    <title>Sample Document</title>
    <style>
        .highlight { color: red; font-weight: bold; }
    </style>
</head>
<body>
    <h1>Chapter 1: Introduction</h1>
    <p>This is the <span class='highlight'>first</span> paragraph of the introduction.</p>
    <p>Another paragraph with <b>bold</b> text.</p>

    <h1>Chapter 2: Details</h1>
    <p>Details start here. <i>Italic text</i> is also present.</p>
    <p>More details in the second chapter.</p>

    <h1>Chapter 3: Conclusion</h1>
    <p>The conclusion wraps up the document.</p>
</body>
</html>";

        // Write HTML to a temporary file
        string htmlPath = Path.Combine(outputDir, "sample.html");
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML document
        Document sourceDoc = new Document(htmlPath);

        // Find all Heading 1 paragraphs (chapters)
        List<Paragraph> chapterHeadings = new List<Paragraph>();
        foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            string styleName = para.ParagraphFormat.Style?.Name;
            if (!string.IsNullOrEmpty(styleName) && styleName.Equals("Heading 1", StringComparison.OrdinalIgnoreCase))
                chapterHeadings.Add(para);
        }

        if (chapterHeadings.Count == 0)
            throw new InvalidOperationException("No chapter headings (Heading 1) were found in the document.");

        // Split each chapter into its own DOCX
        for (int i = 0; i < chapterHeadings.Count; i++)
        {
            Paragraph startHeading = chapterHeadings[i];
            Paragraph endHeading = (i + 1 < chapterHeadings.Count) ? chapterHeadings[i + 1] : null;

            // Create a new empty document
            Document chapterDoc = new Document();
            // Remove the default empty section body content
            chapterDoc.FirstSection.Body.RemoveAllChildren();

            // Import headers and footers from the source document's first section (if any)
            NodeImporter headerFooterImporter = new NodeImporter(sourceDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);
            Section sourceSection = sourceDoc.FirstSection;
            foreach (HeaderFooter headerFooter in sourceSection.HeadersFooters)
            {
                HeaderFooter imported = (HeaderFooter)headerFooterImporter.ImportNode(headerFooter, true);
                chapterDoc.FirstSection.HeadersFooters.Add(imported);
            }

            // Import nodes belonging to this chapter
            NodeImporter importer = new NodeImporter(sourceDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);
            Node currentNode = startHeading;
            while (currentNode != null && currentNode != endHeading)
            {
                Node importedNode = importer.ImportNode(currentNode, true);
                chapterDoc.FirstSection.Body.AppendChild(importedNode);
                currentNode = currentNode.NextSibling;
            }

            // Save the chapter document
            string chapterPath = Path.Combine(outputDir, $"Chapter_{i + 1}.docx");
            chapterDoc.Save(chapterPath);
        }

        // Verify that the expected files were created
        for (int i = 0; i < chapterHeadings.Count; i++)
        {
            string expectedPath = Path.Combine(outputDir, $"Chapter_{i + 1}.docx");
            if (!File.Exists(expectedPath))
                throw new FileNotFoundException($"Expected chapter file not found: {expectedPath}");
        }

        // Optional: clean up temporary HTML file (comment out if you want to keep it)
        // File.Delete(htmlPath);
    }
}
