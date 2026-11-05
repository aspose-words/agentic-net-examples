using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Sample HTML containing three chapters with inline styles.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head><title>Sample Book</title></head>
<body>
<h1>Chapter 1</h1>
<p>This is the <span style='color:red;'>first</span> paragraph of chapter one.</p>
<p>Another paragraph with <b>bold</b> text.</p>
<h1>Chapter 2</h1>
<p>Content of <i>chapter</i> two starts here.</p>
<p style='font-size:14pt;'>A paragraph with larger font.</p>
<h1>Chapter 3</h1>
<p>Final chapter with a <span style='background-color:yellow;'>highlighted</span> word.</p>
</body>
</html>";

        // Write the HTML to a temporary file so Aspose.Words can load it.
        string sourceHtmlPath = "source.html";
        File.WriteAllText(sourceHtmlPath, htmlContent);

        // Load the HTML document.
        Document sourceDoc = new Document(sourceHtmlPath);

        // Locate all Heading 1 paragraphs – they mark the start of each chapter.
        var headingParagraphs = sourceDoc
            .GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Where(p => p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            .ToList();

        if (headingParagraphs.Count == 0)
            throw new InvalidOperationException("No chapters (Heading 1) found in the source document.");

        // Split each chapter into its own DOCX file.
        for (int i = 0; i < headingParagraphs.Count; i++)
        {
            Paragraph startHeading = headingParagraphs[i];
            Paragraph endHeading = (i + 1 < headingParagraphs.Count) ? headingParagraphs[i + 1] : null;

            // Create a new empty document that will hold the chapter.
            Document chapterDoc = new Document(); // Contains a default section.
            Body chapterBody = chapterDoc.FirstSection.Body;

            // Importer to copy nodes while preserving source formatting (including inline styles).
            NodeImporter importer = new NodeImporter(sourceDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);

            // Copy nodes from the start heading up to (but not including) the next heading.
            Node currentNode = startHeading;
            while (currentNode != null && currentNode != endHeading)
            {
                Node importedNode = importer.ImportNode(currentNode, true);
                chapterBody.AppendChild(importedNode);
                currentNode = currentNode.NextSibling;
            }

            // Save the chapter as a separate DOCX file.
            string chapterFileName = $"Chapter_{i + 1}.docx";
            chapterDoc.Save(chapterFileName, SaveFormat.Docx);

            // Verify that the file was created.
            if (!File.Exists(chapterFileName))
                throw new InvalidOperationException($"Failed to create split file: {chapterFileName}");
        }

        // Clean up the temporary HTML file.
        if (File.Exists(sourceHtmlPath))
            File.Delete(sourceHtmlPath);
    }
}
