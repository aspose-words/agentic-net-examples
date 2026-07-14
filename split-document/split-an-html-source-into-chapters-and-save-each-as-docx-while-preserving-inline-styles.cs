using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitDemo");
        string inputDir = Path.Combine(baseDir, "Input");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample HTML file containing heading paragraphs (chapters) and styled text.
        string htmlPath = Path.Combine(inputDir, "Sample.html");
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head><title>Sample</title></head>
<body>
<h1>Chapter 1</h1>
<p style='color:red;'>This is <b>red</b> text in chapter 1.</p>
<p>This paragraph continues chapter 1.</p>
<h1>Chapter 2</h1>
<p style='font-size:14pt;'>This is a paragraph in <i>chapter 2</i> with larger font.</p>
<p style='background-color:yellow;'>Highlighted text in chapter 2.</p>
<h1>Chapter 3</h1>
<p>Final chapter without extra styling.</p>
</body>
</html>";
        File.WriteAllText(htmlPath, htmlContent);

        // Load the HTML document.
        Document srcDoc = new Document(htmlPath);

        // Identify all paragraphs that use the Heading 1 style – they will serve as chapter delimiters.
        var headingParagraphs = srcDoc.GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .Where(p => p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            .ToList();

        if (!headingParagraphs.Any())
            throw new InvalidOperationException("No heading paragraphs found to split the document.");

        int chapterIndex = 0;

        // Iterate over each heading and extract the content up to the next heading.
        for (int i = 0; i < headingParagraphs.Count; i++)
        {
            Paragraph startHeading = headingParagraphs[i];
            chapterIndex++;

            // Create a new empty document for the current chapter.
            Document chapterDoc = new Document();
            chapterDoc.RemoveAllChildren();

            // Add a fresh section with a body.
            Section chapterSection = new Section(chapterDoc);
            chapterDoc.AppendChild(chapterSection);
            Body chapterBody = new Body(chapterDoc);
            chapterSection.AppendChild(chapterBody);

            // Use a NodeImporter to copy nodes while preserving source formatting.
            NodeImporter importer = new NodeImporter(srcDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);

            // Import the heading itself.
            Node importedHeading = importer.ImportNode(startHeading, true);
            chapterBody.AppendChild(importedHeading);

            // Walk through subsequent nodes until the next Heading 1 paragraph or the end of the document.
            Node currentNode = startHeading.NextSibling;
            while (currentNode != null &&
                   !(currentNode.NodeType == NodeType.Paragraph &&
                     ((Paragraph)currentNode).ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1))
            {
                Node importedNode = importer.ImportNode(currentNode, true);
                chapterBody.AppendChild(importedNode);
                currentNode = currentNode.NextSibling;
            }

            // Save the chapter as a DOCX file.
            string chapterPath = Path.Combine(outputDir, $"Chapter_{chapterIndex}.docx");
            chapterDoc.Save(chapterPath);
        }

        // Verify that the expected number of chapter files were created.
        int expectedFiles = headingParagraphs.Count;
        int actualFiles = Directory.GetFiles(outputDir, "*.docx").Length;
        if (actualFiles != expectedFiles)
            throw new InvalidOperationException($"Expected {expectedFiles} chapter files, but found {actualFiles}.");
    }
}
