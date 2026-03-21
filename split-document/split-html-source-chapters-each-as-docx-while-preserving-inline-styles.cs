using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

namespace HtmlChapterSplitter
{
    class Program
    {
        static void Main()
        {
            // Sample HTML content.
            const string htmlContent = @"
                <html>
                <body>
                    <h1>Chapter 1</h1>
                    <p>This is the first paragraph of chapter 1.</p>
                    <p>Another paragraph in chapter 1.</p>
                    <h2>Section 1.1</h2>
                    <p>Content of section 1.1.</p>
                    <h1>Chapter 2</h1>
                    <p>First paragraph of chapter 2.</p>
                </body>
                </html>";

            // Create a temporary folder for output.
            string outputFolder = Path.Combine(Path.GetTempPath(), "HtmlChapterSplitter", "Chapters");
            Directory.CreateDirectory(outputFolder);

            // Load the HTML document from the string.
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent));
            var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
            Document sourceDoc = new Document(stream, loadOptions);

            // List to hold each chapter document.
            var chapterDocs = new List<Document>();

            Document currentChapter = null;
            NodeImporter importer = null;

            // Iterate through all paragraphs in the source document in order.
            foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Determine if the paragraph is a heading (any heading style).
                bool isHeading = para.ParagraphFormat.IsHeading;

                // If we encounter a heading, start a new chapter.
                if (isHeading)
                {
                    currentChapter = new Document();
                    currentChapter.EnsureMinimum();
                    importer = new NodeImporter(sourceDoc, currentChapter, ImportFormatMode.KeepSourceFormatting);
                    chapterDocs.Add(currentChapter);
                }

                // If we have not yet started a chapter (e.g., content before first heading),
                // create a default chapter.
                if (currentChapter == null)
                {
                    currentChapter = new Document();
                    currentChapter.EnsureMinimum();
                    importer = new NodeImporter(sourceDoc, currentChapter, ImportFormatMode.KeepSourceFormatting);
                    chapterDocs.Add(currentChapter);
                }

                // Import the paragraph (and its child nodes) into the current chapter.
                Node importedPara = importer.ImportNode(para, true);
                currentChapter.FirstSection.Body.AppendChild(importedPara);
            }

            // Save each chapter as a separate DOCX file.
            for (int i = 0; i < chapterDocs.Count; i++)
            {
                string chapterFileName = Path.Combine(outputFolder, $"Chapter_{i + 1}.docx");
                chapterDocs[i].Save(chapterFileName);
            }

            Console.WriteLine("Splitting complete. Chapters saved to: " + outputFolder);
        }
    }
}
