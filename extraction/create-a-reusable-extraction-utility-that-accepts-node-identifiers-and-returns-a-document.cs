using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ExtractionExample
{
    // Reusable extraction utility.
    public static class ExtractionUtility
    {
        // Extracts the content of a bookmark into a new document while preserving formatting.
        public static Document ExtractByBookmark(Document source, string bookmarkName)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(bookmarkName)) throw new ArgumentException("Bookmark name must be provided.", nameof(bookmarkName));

            Bookmark bookmark = source.Range.Bookmarks[bookmarkName];
            if (bookmark == null)
                throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

            // Prepare the destination document.
            Document result = new Document();
            result.RemoveAllChildren();

            Section section = new Section(result);
            result.AppendChild(section);
            Body body = new Body(result);
            section.AppendChild(body);

            // Importer to copy nodes from source to result.
            NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);

            // Walk through nodes from the start of the bookmark to its end.
            Node startNode = bookmark.BookmarkStart;
            Node endNode = bookmark.BookmarkEnd;
            Node current = startNode;

            while (current != null)
            {
                // Skip bookmark start/end markers themselves.
                if (current.NodeType != NodeType.BookmarkStart && current.NodeType != NodeType.BookmarkEnd)
                {
                    Node imported = importer.ImportNode(current, true);

                    // Only block‑level nodes can be appended directly to the body.
                    if (imported.NodeType == NodeType.Paragraph || imported.NodeType == NodeType.Table)
                    {
                        body.AppendChild(imported);
                    }
                    else if (imported.NodeType == NodeType.Run)
                    {
                        // Wrap inline runs in a paragraph.
                        Paragraph para = new Paragraph(result);
                        para.AppendChild(imported);
                        body.AppendChild(para);
                    }
                }

                if (current == endNode)
                    break;

                current = current.NextSibling;
            }

            // Fallback: if nothing was added, write the plain bookmark text.
            if (!body.HasChildNodes)
            {
                Paragraph para = new Paragraph(result);
                para.AppendChild(new Run(result, bookmark.Text));
                body.AppendChild(para);
            }

            return result;
        }

        // Extracts a range of paragraphs (by zero‑based indices, inclusive) into a new document.
        public static Document ExtractParagraphRange(Document source, int startIndex, int endIndex)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (startIndex < 0 || endIndex < 0)
                throw new ArgumentException("Paragraph indices must be non‑negative.");
            if (endIndex < startIndex)
                throw new ArgumentException("End index must be greater than or equal to start index.");

            ParagraphCollection paragraphs = source.FirstSection.Body.Paragraphs;
            if (startIndex >= paragraphs.Count || endIndex >= paragraphs.Count)
                throw new InvalidOperationException("Paragraph index out of range.");

            Document result = new Document();
            result.RemoveAllChildren();

            Section section = new Section(result);
            result.AppendChild(section);
            Body body = new Body(result);
            section.AppendChild(body);

            // Use an importer to preserve formatting when moving nodes between documents.
            NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);

            for (int i = startIndex; i <= endIndex; i++)
            {
                Paragraph srcPara = paragraphs[i];
                if (srcPara == null)
                    throw new InvalidOperationException($"Paragraph at index {i} is null.");

                Node imported = importer.ImportNode(srcPara, true);
                body.AppendChild(imported);
            }

            return result;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample source document.
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            builder.Writeln("Paragraph 0 - Intro");
            builder.Writeln("Paragraph 1 - Before bookmark");
            builder.StartBookmark("SampleBookmark");
            builder.Writeln("This is the bookmarked paragraph.");
            builder.EndBookmark("SampleBookmark");
            builder.Writeln("Paragraph 3 - After bookmark");
            builder.Writeln("Paragraph 4 - Additional content");

            // Save the source document.
            const string sourcePath = "source.docx";
            sourceDoc.Save(sourcePath);

            // Load the source document.
            Document loadedDoc = new Document(sourcePath);

            // Use the extraction utility to extract the bookmark content.
            Document bookmarkExtracted = ExtractionUtility.ExtractByBookmark(loadedDoc, "SampleBookmark");
            const string bookmarkOutput = "bookmark-extracted.docx";
            bookmarkExtracted.Save(bookmarkOutput);

            // Validate bookmark extraction output.
            if (!File.Exists(bookmarkOutput))
                throw new InvalidOperationException("Bookmark extraction output was not created.");

            // Extract paragraphs 1 through 3 (inclusive).
            Document paragraphExtracted = ExtractionUtility.ExtractParagraphRange(loadedDoc, 1, 3);
            const string paragraphOutput = "paragraphs-extracted.docx";
            paragraphExtracted.Save(paragraphOutput);

            // Validate paragraph extraction output.
            if (!File.Exists(paragraphOutput))
                throw new InvalidOperationException("Paragraph range extraction output was not created.");

            // Example completed successfully.
        }
    }
}
