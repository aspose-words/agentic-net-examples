using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

namespace SplitEpubIntoHtmlChapters
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample EPUB with three chapters.
            string epubPath = "sample.epub";
            CreateSampleEpub(epubPath);

            // Load the EPUB document.
            Document epubDoc = new Document(epubPath);

            // Split the document into chapters based on Heading 1 paragraphs.
            List<Document> chapterDocs = SplitIntoChaptersByHeading(epubDoc, StyleIdentifier.Heading1);

            // Save each chapter as an individual HTML file.
            for (int i = 0; i < chapterDocs.Count; i++)
            {
                string htmlFile = $"Chapter_{i + 1}.html";
                chapterDocs[i].Save(htmlFile, SaveFormat.Html);
                if (!File.Exists(htmlFile))
                    throw new InvalidOperationException($"Failed to create {htmlFile}");
            }

            // Verify that the expected number of HTML files were created.
            int expectedCount = chapterDocs.Count;
            int actualCount = Directory.GetFiles(Directory.GetCurrentDirectory(), "Chapter_*.html").Length;
            if (actualCount != expectedCount)
                throw new InvalidOperationException($"Expected {expectedCount} HTML files, but found {actualCount}.");
        }

        private static void CreateSampleEpub(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int chapter = 1; chapter <= 3; chapter++)
            {
                // Heading for the chapter.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                builder.Writeln($"Chapter {chapter}");

                // Normal content paragraphs.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                builder.Writeln($"This is the first paragraph of chapter {chapter}.");
                builder.Writeln($"This is the second paragraph of chapter {chapter}.");
            }

            doc.Save(path, SaveFormat.Epub);
            if (!File.Exists(path))
                throw new InvalidOperationException($"Failed to create sample EPUB at {path}");
        }

        private static List<Document> SplitIntoChaptersByHeading(Document source, StyleIdentifier headingStyle)
        {
            List<Document> chapters = new List<Document>();
            Document currentChapter = null;

            // Retrieve all paragraphs in document order.
            NodeCollection paragraphs = source.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph para in paragraphs)
            {
                if (para.ParagraphFormat.StyleIdentifier == headingStyle)
                {
                    // Start a new chapter document.
                    currentChapter = new Document();
                    // Import the heading paragraph into the new document.
                    Node importedHeading = currentChapter.ImportNode(para, true);
                    currentChapter.FirstSection.Body.AppendChild(importedHeading);
                    chapters.Add(currentChapter);
                }
                else if (currentChapter != null)
                {
                    // Import regular paragraph into the current chapter.
                    Node importedPara = currentChapter.ImportNode(para, true);
                    currentChapter.FirstSection.Body.AppendChild(importedPara);
                }
            }

            if (chapters.Count == 0)
                throw new InvalidOperationException("No chapters were found based on the specified heading style.");

            return chapters;
        }
    }
}
