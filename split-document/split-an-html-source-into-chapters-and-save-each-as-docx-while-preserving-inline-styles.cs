using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(outputDir);

        // Sample HTML source containing headings and inline styles.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head><title>Sample Document</title></head>
<body>
<h1 style='color:blue;'>Chapter 1</h1>
<p style='font-weight:bold;'>This is the first paragraph of chapter 1.</p>
<p>This is a normal paragraph.</p>
<h1 style='color:green;'>Chapter 2</h1>
<p style='font-style:italic;'>First paragraph of chapter 2 with italic style.</p>
<p>Another paragraph in chapter 2.</p>
<h1 style='color:red;'>Chapter 3</h1>
<p>Content of chapter 3.</p>
</body>
</html>";

        // Load the HTML into an Aspose.Words Document using a MemoryStream.
        using (MemoryStream htmlStream = new MemoryStream())
        using (StreamWriter writer = new StreamWriter(htmlStream))
        {
            writer.Write(htmlContent);
            writer.Flush();
            htmlStream.Position = 0;

            LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
            Document sourceDoc = new Document(htmlStream, loadOptions);

            // Collect all paragraphs to identify heading paragraphs (Heading 1).
            NodeCollection paragraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);

            List<Document> chapterDocs = new List<Document>();
            Document currentChapter = null;
            NodeImporter importer = null;
            int chapterIndex = 0;

            foreach (Paragraph para in paragraphs)
            {
                // Determine if the paragraph is a Heading 1.
                bool isHeading1 = para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1;

                if (isHeading1)
                {
                    // When a new heading is found, finalize the previous chapter (if any).
                    if (currentChapter != null)
                    {
                        string chapterPath = Path.Combine(outputDir, $"Chapter_{chapterIndex}.docx");
                        currentChapter.Save(chapterPath);
                    }

                    // Start a new chapter document.
                    chapterIndex++;
                    currentChapter = new Document();
                    // Ensure the document has a section and body.
                    currentChapter.RemoveAllChildren();
                    Section sec = new Section(currentChapter);
                    currentChapter.AppendChild(sec);
                    Body body = new Body(currentChapter);
                    sec.AppendChild(body);

                    // Prepare an importer for this chapter.
                    importer = new NodeImporter(sourceDoc, currentChapter, ImportFormatMode.KeepSourceFormatting);
                }

                // If we have an active chapter, import the current paragraph.
                if (currentChapter != null && importer != null)
                {
                    Node importedNode = importer.ImportNode(para, true);
                    currentChapter.FirstSection.Body.AppendChild(importedNode);
                }
            }

            // Save the last chapter if it exists.
            if (currentChapter != null)
            {
                string chapterPath = Path.Combine(outputDir, $"Chapter_{chapterIndex}.docx");
                currentChapter.Save(chapterPath);
            }

            // Validation: ensure that at least three chapter files were created.
            string[] savedFiles = Directory.GetFiles(outputDir, "Chapter_*.docx");
            if (savedFiles.Length < 3)
                throw new InvalidOperationException("Expected at least three chapter files to be created.");
        }
    }
}
