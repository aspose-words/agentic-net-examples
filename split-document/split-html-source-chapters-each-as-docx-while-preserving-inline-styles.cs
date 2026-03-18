using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

namespace HtmlChapterSplitter
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file.
            string htmlPath = @"C:\Input\source.html";

            // Folder where the resulting DOCX chapters will be saved.
            string outputFolder = @"C:\Output\Chapters";

            // Ensure the output folder exists.
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Load the HTML document.
            Document sourceDoc = new Document(htmlPath);

            // List to hold each chapter document.
            List<Document> chapterDocs = new List<Document>();

            // The current chapter document we are building.
            Document currentChapter = null;

            // NodeImporter to efficiently import nodes from the source document.
            NodeImporter importer = null;

            // Iterate through all paragraphs in the source document in order.
            foreach (Paragraph para in sourceDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Determine if the paragraph is a heading (any heading style).
                bool isHeading = para.ParagraphFormat.IsHeading;

                // If we encounter a heading, start a new chapter.
                if (isHeading)
                {
                    // Finish the previous chapter (if any) and start a new one.
                    currentChapter = new Document();
                    currentChapter.EnsureMinimum();

                    // Create a fresh importer for this chapter.
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
