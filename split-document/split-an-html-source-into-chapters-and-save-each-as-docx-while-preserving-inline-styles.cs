using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Sample HTML with headings and inline styles
        string html = @"
        <html>
        <body>
            <h1>Chapter One</h1>
            <p>This is the <span style='color:red;'>first</span> paragraph of chapter one.</p>
            <p>Another paragraph with <b>bold</b> text.</p>
            <h1>Chapter Two</h1>
            <p>Content of <i>chapter two</i> starts here.</p>
            <p>Second paragraph with <span style='font-size:20px;'>large text</span>.</p>
            <h1>Chapter Three</h1>
            <p>Final chapter content.</p>
        </body>
        </html>";

        // Write HTML to a temporary file
        string htmlPath = Path.Combine(Path.GetTempPath(), "sample.html");
        File.WriteAllText(htmlPath, html);

        // Load HTML into Aspose.Words Document
        Document srcDoc = new Document(htmlPath);

        // Find all Heading 1 paragraphs (chapters)
        List<Paragraph> headings = new List<Paragraph>();
        foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                headings.Add(para);
        }

        if (headings.Count == 0)
            throw new Exception("No chapter headings found in the source document.");

        // Split document by headings
        for (int i = 0; i < headings.Count; i++)
        {
            Paragraph startHeading = headings[i];
            Node endNode = (i + 1 < headings.Count) ? (Node)headings[i + 1] : null;

            // Create a new empty document for the chapter
            Document chapterDoc = new Document();
            Section chapterSection = chapterDoc.FirstSection;
            Body chapterBody = chapterSection.Body;

            // Copy nodes from startHeading up to (but not including) endNode
            Node current = startHeading;
            while (current != null && current != endNode)
            {
                Node imported = chapterDoc.ImportNode(current, true);
                chapterBody.AppendChild(imported);
                current = current.NextSibling;
            }

            // Save the chapter as DOCX
            string chapterFile = Path.Combine(Directory.GetCurrentDirectory(), $"Chapter_{i + 1}.docx");
            chapterDoc.Save(chapterFile, SaveFormat.Docx);

            // Verify the file was created
            if (!File.Exists(chapterFile))
                throw new Exception($"Failed to create chapter file: {chapterFile}");
        }

        // Clean up temporary HTML file
        File.Delete(htmlPath);
    }
}
