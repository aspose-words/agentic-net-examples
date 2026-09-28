using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ExtractionExample
{
    // Reusable extraction utility.
    public static class ExtractionUtility
    {
        // Extracts the content of a bookmark (including the paragraph that contains it) into a new document.
        public static Document ExtractByBookmark(Document source, string bookmarkName)
        {
            Bookmark bookmark = source.Range.Bookmarks[bookmarkName];
            if (bookmark == null)
                throw new InvalidOperationException($"Bookmark '{bookmarkName}' not found.");

            // The bookmark is located inside a paragraph. Get that paragraph.
            Paragraph containingParagraph = bookmark.BookmarkStart.ParentNode as Paragraph;
            if (containingParagraph == null)
                throw new InvalidOperationException("Bookmark is not inside a paragraph.");

            Document result = CreateEmptyDocument();

            // Import the paragraph into the new document.
            NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);
            Node importedParagraph = importer.ImportNode(containingParagraph, true);
            result.FirstSection.Body.AppendChild(importedParagraph);

            return result;
        }

        // Extracts a paragraph by zero‑based index from the document body.
        public static Document ExtractParagraphByIndex(Document source, int paragraphIndex)
        {
            Paragraph paragraph = source.FirstSection.Body.Paragraphs[paragraphIndex];
            if (paragraph == null)
                throw new InvalidOperationException($"Paragraph at index {paragraphIndex} not found.");

            Document result = CreateEmptyDocument();

            // Import the paragraph into the new document.
            NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);
            Node importedParagraph = importer.ImportNode(paragraph, true);
            result.FirstSection.Body.AppendChild(importedParagraph);

            return result;
        }

        // Extracts a table by zero‑based index from the document.
        public static Document ExtractTableByIndex(Document source, int tableIndex)
        {
            NodeCollection tables = source.GetChildNodes(NodeType.Table, true);
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new InvalidOperationException($"Table at index {tableIndex} not found.");

            Table table = tables[tableIndex] as Table;
            if (table == null)
                throw new InvalidOperationException($"Table at index {tableIndex} not found.");

            Document result = CreateEmptyDocument();

            // Import the table into the new document.
            NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(table, true);
            result.FirstSection.Body.AppendChild(importedTable);

            return result;
        }

        // Helper that creates a new empty document with a single section and body.
        private static Document CreateEmptyDocument()
        {
            Document doc = new Document();
            doc.RemoveAllChildren();

            Section section = new Section(doc);
            doc.AppendChild(section);

            Body body = new Body(doc);
            section.AppendChild(body);

            return doc;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample source document.
            Document source = new Document();
            DocumentBuilder builder = new DocumentBuilder(source);

            builder.Writeln("Paragraph before bookmark.");
            builder.StartBookmark("SampleBookmark");
            builder.Writeln("This is the bookmarked paragraph.");
            builder.EndBookmark("SampleBookmark");
            builder.Writeln("Paragraph after bookmark.");

            // Add a second paragraph for index extraction.
            builder.Writeln("Second paragraph for index extraction.");

            // Add a sample table.
            builder.StartTable();
            builder.InsertCell(); builder.Write("Header 1");
            builder.InsertCell(); builder.Write("Header 2");
            builder.EndRow();
            builder.InsertCell(); builder.Write("Row1 Cell1");
            builder.InsertCell(); builder.Write("Row1 Cell2");
            builder.EndRow();
            builder.EndTable();

            // Save the source document.
            const string sourcePath = "source.docx";
            source.Save(sourcePath);

            // Load the document (simulating a real scenario).
            Document loaded = new Document(sourcePath);

            // 1. Extract by bookmark.
            Document bookmarkExtract = ExtractionUtility.ExtractByBookmark(loaded, "SampleBookmark");
            const string bookmarkOutput = "bookmark-extract.docx";
            bookmarkExtract.Save(bookmarkOutput);
            if (!File.Exists(bookmarkOutput))
                throw new InvalidOperationException("Bookmark extraction output was not created.");

            // 2. Extract the second paragraph (index 1).
            Document paragraphExtract = ExtractionUtility.ExtractParagraphByIndex(loaded, 1);
            const string paragraphOutput = "paragraph-extract.docx";
            paragraphExtract.Save(paragraphOutput);
            if (!File.Exists(paragraphOutput))
                throw new InvalidOperationException("Paragraph extraction output was not created.");

            // 3. Extract the first table (index 0).
            Document tableExtract = ExtractionUtility.ExtractTableByIndex(loaded, 0);
            const string tableOutput = "table-extract.docx";
            tableExtract.Save(tableOutput);
            if (!File.Exists(tableOutput))
                throw new InvalidOperationException("Table extraction output was not created.");

            // All extractions succeeded.
            Console.WriteLine("Extraction completed successfully.");
        }
    }
}
