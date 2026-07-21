using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace AsposeWordsTableTocExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Table of Contents (TOC) field at the beginning of the document.
            // Use a simple switch string; the field will be updated later.
            Field tocField = builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            // Cast to FieldToc to set additional properties.
            if (tocField is FieldToc toc)
            {
                // The TOC will include only entries that are inside the bookmark named "TableBookmark".
                toc.BookmarkName = "TableBookmark";
                // Make TOC entries clickable hyperlinks.
                toc.InsertHyperlinks = true;
            }

            // Insert a page break so the table appears on a new page.
            builder.InsertBreak(BreakType.PageBreak);

            // Start a bookmark that will surround the table and its heading.
            builder.StartBookmark("TableBookmark");

            // Insert a heading that will appear in the TOC as the table entry.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Table 1: Sample Data");

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // End the bookmark.
            builder.EndBookmark("TableBookmark");

            // Update all fields (including the TOC) so that the entry appears.
            doc.UpdateFields();

            // Define the output path relative to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithToc.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"Document saved successfully to: {outputPath}");
            }
            else
            {
                throw new InvalidOperationException("Failed to save the document.");
            }
        }
    }
}
