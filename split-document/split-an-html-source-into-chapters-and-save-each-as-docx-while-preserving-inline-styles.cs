using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Sample HTML containing two chapters with inline styles.
        string html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Sample Book</title>
</head>
<body>
    <h1>Chapter 1: Introduction</h1>
    <p style='color:red;'>This is the first paragraph of chapter 1 with red text.</p>
    <p>This is the second paragraph of chapter 1.</p>
    <h1>Chapter 2: Advanced Topics</h1>
    <p style='font-weight:bold;'>Bold paragraph in chapter 2.</p>
    <p>Regular paragraph in chapter 2.</p>
</body>
</html>";

        // Load the HTML into an Aspose.Words Document from a memory stream.
        using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
        {
            // Reset the stream position before loading.
            htmlStream.Position = 0;

            LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
            Document sourceDoc = new Document(htmlStream, loadOptions);

            // Locate all Heading 1 paragraphs – they mark the start of each chapter.
            List<Paragraph> chapterHeadings = new List<Paragraph>();
            foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Aspose.Words maps <h1> to the style named "Heading 1".
                if (para.ParagraphFormat.Style != null && para.ParagraphFormat.Style.Name == "Heading 1")
                {
                    chapterHeadings.Add(para);
                }
            }

            if (chapterHeadings.Count == 0)
                throw new InvalidOperationException("No chapter headings (Heading 1) were found in the source document.");

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Split each chapter into a separate DOCX file.
            for (int i = 0; i < chapterHeadings.Count; i++)
            {
                Paragraph startHeading = chapterHeadings[i];
                Paragraph endHeading = (i + 1 < chapterHeadings.Count) ? chapterHeadings[i + 1] : null;

                // Create a new empty document for the chapter.
                Document chapterDoc = new Document();

                // Importer to copy nodes while preserving source formatting.
                NodeImporter importer = new NodeImporter(sourceDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);

                // Import the first (empty) section from the source to preserve section settings.
                Section importedSection = (Section)chapterDoc.ImportNode(sourceDoc.FirstSection, true);
                chapterDoc.RemoveAllChildren(); // Remove the default empty section.
                chapterDoc.AppendChild(importedSection);
                Body chapterBody = importedSection.Body;
                chapterBody.RemoveAllChildren(); // Start with an empty body.

                // Copy nodes from the start heading up to (but not including) the next heading.
                Node curNode = startHeading;
                while (curNode != null && curNode != endHeading)
                {
                    Node importedNode = importer.ImportNode(curNode, true);
                    chapterBody.AppendChild(importedNode);
                    curNode = curNode.NextSibling;
                }

                // Save the chapter document.
                string chapterPath = Path.Combine(outputDir, $"Chapter_{i + 1}.docx");
                chapterDoc.Save(chapterPath, SaveFormat.Docx);

                // Verify that the file was created.
                if (!File.Exists(chapterPath))
                    throw new InvalidOperationException($"Failed to create split file: {chapterPath}");
            }
        }

        // All chapters have been split and saved successfully.
        Console.WriteLine("Chapter splitting completed.");
    }
}
