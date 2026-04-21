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

            // Insert a Table of Contents field at the beginning of the document.
            // The \\b switch limits the TOC to entries that appear inside the bookmark named "TableBookmark".
            // The \\h switch makes the entries hyperlinks.
            // The \\z switch hides page numbers in web layout.
            // The \\u switch uses outline levels.
            Field tocField = builder.InsertTableOfContents("\\b TableBookmark \\h \\z \\u");
            // Cast to FieldToc to set the bookmark name explicitly (optional, already set via switch).
            if (tocField is FieldToc toc)
            {
                toc.BookmarkName = "TableBookmark";
                toc.InsertHyperlinks = true;
            }

            // Add a page break after the TOC so the table appears on a new page.
            builder.InsertBreak(BreakType.PageBreak);

            // Start a bookmark that will contain the table and its title.
            builder.StartBookmark("TableBookmark");

            // Insert a heading that will be captured by the TOC.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Table 1: Sample Data");

            // Build a simple 2x2 table.
            builder.StartTable();

            // First row – header cells.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row – data cells.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // End the bookmark.
            builder.EndBookmark("TableBookmark");

            // Update all fields (including the TOC) so that the entry appears.
            doc.UpdateFields();

            // Define the output file path.
            string outputPath = "TableOfContentsWithTable.docx";

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
