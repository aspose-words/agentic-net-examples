using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Sample HTML containing two chapters.
        string html = @"
        <html>
            <body>
                <h1>Chapter 1</h1>
                <p style='color:red; font-weight:bold;'>This is the first paragraph of chapter 1.</p>
                <p>This is another paragraph of chapter 1.</p>
                <h1>Chapter 2</h1>
                <p style='font-style:italic;'>This is the first paragraph of chapter 2.</p>
                <p>This is another paragraph of chapter 2.</p>
            </body>
        </html>";

        // Load the HTML into an Aspose.Words Document.
        using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
        {
            Document sourceDoc = new Document(htmlStream);

            // Locate all Heading 1 paragraphs (chapter titles).
            List<Paragraph> chapterHeadings = new List<Paragraph>();
            foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                    chapterHeadings.Add(para);
            }

            if (chapterHeadings.Count == 0)
                throw new Exception("No Heading 1 paragraphs were found in the source document.");

            // Split the document by each heading.
            for (int i = 0; i < chapterHeadings.Count; i++)
            {
                Document chapterDoc = new Document(); // New empty document.
                NodeImporter importer = new NodeImporter(sourceDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);

                Paragraph startHeading = chapterHeadings[i];
                Node nextHeading = (i + 1 < chapterHeadings.Count) ? (Node)chapterHeadings[i + 1] : null;

                // Copy nodes from the start heading up to (but not including) the next heading.
                Node currentNode = startHeading;
                while (currentNode != null && currentNode != nextHeading)
                {
                    Node importedNode = importer.ImportNode(currentNode, true);
                    chapterDoc.FirstSection.Body.AppendChild(importedNode);
                    currentNode = currentNode.NextSibling;
                }

                // Save the chapter as a DOCX file.
                string fileName = $"Chapter_{i + 1}.docx";
                chapterDoc.Save(fileName);

                // Verify the file was created.
                if (!File.Exists(fileName))
                    throw new Exception($"Failed to create output file: {fileName}");
            }
        }

        Console.WriteLine("Document split completed. Chapter files created.");
    }
}
