using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create sample HTML with three chapters and inline styles.
        string html = @"
<html>
<body>
<h1>Chapter 1</h1>
<p style='color:red;'>This is the first paragraph of chapter 1.</p>
<p>This is another paragraph.</p>
<h1>Chapter 2</h1>
<p style='font-weight:bold;'>Bold paragraph in chapter 2.</p>
<p>Normal paragraph.</p>
<h1>Chapter 3</h1>
<p>Last chapter content.</p>
</body>
</html>";

        // Write the HTML to a temporary file.
        string htmlPath = "sample.html";
        File.WriteAllText(htmlPath, html);

        // Load the HTML into an Aspose.Words Document.
        Document sourceDoc = new Document(htmlPath);

        // Find all heading paragraphs (assumed to be Heading 1 style).
        List<Paragraph> headings = new List<Paragraph>();
        foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleName == "Heading 1")
                headings.Add(para);
        }

        if (headings.Count == 0)
            throw new Exception("No headings found to split the document.");

        // Split the document into chapters based on headings.
        for (int i = 0; i < headings.Count; i++)
        {
            Paragraph startHeading = headings[i];
            Node endNode = (i + 1 < headings.Count) ? (Node)headings[i + 1] : null;

            // Create a new empty document for the chapter.
            Document chapterDoc = new Document();
            // Remove the default empty paragraph that Aspose.Words adds.
            chapterDoc.FirstSection.Body.RemoveAllChildren();

            NodeImporter importer = new NodeImporter(sourceDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);
            Node currentNode = startHeading;

            while (currentNode != null && currentNode != endNode)
            {
                Node importedNode = importer.ImportNode(currentNode, true);
                chapterDoc.FirstSection.Body.AppendChild(importedNode);
                currentNode = currentNode.NextSibling;
            }

            string chapterFileName = $"Chapter_{i + 1}.docx";
            chapterDoc.Save(chapterFileName);

            // Validate that the file was created.
            if (!File.Exists(chapterFileName))
                throw new Exception($"Failed to create split file: {chapterFileName}");
        }

        // Clean up temporary HTML file.
        if (File.Exists(htmlPath))
            File.Delete(htmlPath);
    }
}
