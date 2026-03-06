using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Notes; // Required for Footnote class

class ExtractContent
{
    static void Main()
    {
        // Load the DOCX document from file system.
        // Uses the Document(string) constructor – the approved loading rule.
        string inputPath = @"C:\Docs\InputDocument.docx";
        Document doc = new Document(inputPath);

        // --------------------------------------------------------------------
        // Example: Extract text between two bookmarks named "Start" and "End".
        // The bookmarks must exist in the document.
        // --------------------------------------------------------------------
        string betweenBookmarks = string.Empty;
        Bookmark startBookmark = doc.Range.Bookmarks["Start"];
        Bookmark endBookmark   = doc.Range.Bookmarks["End"];
        if (startBookmark != null && endBookmark != null)
        {
            // Walk the node tree from the start bookmark to the end bookmark
            // and concatenate the text of Run nodes.
            Node startNode = startBookmark.BookmarkStart;
            Node endNode   = endBookmark.BookmarkEnd;
            StringBuilder sb = new StringBuilder();
            for (Node cur = startNode; cur != null && cur != endNode; cur = cur.NextSibling)
            {
                if (cur.NodeType == NodeType.Run)
                {
                    sb.Append(((Run)cur).Text);
                }
                else if (cur.NodeType == NodeType.Paragraph && cur != startNode)
                {
                    // Preserve paragraph breaks.
                    sb.AppendLine();
                }
            }
            betweenBookmarks = sb.ToString();
        }

        // --------------------------------------------------------------------
        // Collect header and footer text from every section.
        // --------------------------------------------------------------------
        StringBuilder headerFooterBuilder = new StringBuilder();

        foreach (Section section in doc.Sections)
        {
            // Primary header (odd pages) and even header.
            HeaderFooter primaryHeader = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            HeaderFooter evenHeader    = section.HeadersFooters[HeaderFooterType.HeaderEven];
            HeaderFooter firstHeader   = section.HeadersFooters[HeaderFooterType.HeaderFirst];

            // Primary footer (odd pages) and even footer.
            HeaderFooter primaryFooter = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            HeaderFooter evenFooter    = section.HeadersFooters[HeaderFooterType.FooterEven];
            HeaderFooter firstFooter   = section.HeadersFooters[HeaderFooterType.FooterFirst];

            // Append non‑null header/footer texts.
            AppendIfNotNull(primaryHeader, headerFooterBuilder);
            AppendIfNotNull(evenHeader,    headerFooterBuilder);
            AppendIfNotNull(firstHeader,   headerFooterBuilder);
            AppendIfNotNull(primaryFooter, headerFooterBuilder);
            AppendIfNotNull(evenFooter,    headerFooterBuilder);
            AppendIfNotNull(firstFooter,   headerFooterBuilder);
        }

        // --------------------------------------------------------------------
        // Collect all footnote texts.
        // --------------------------------------------------------------------
        StringBuilder footnoteBuilder = new StringBuilder();
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnotes)
        {
            // Footnote.GetText() returns the full footnote text.
            footnoteBuilder.AppendLine(footnote.GetText().Trim());
        }

        // --------------------------------------------------------------------
        // Combine all extracted parts.
        // --------------------------------------------------------------------
        StringBuilder resultBuilder = new StringBuilder();

        resultBuilder.AppendLine("=== Text Between Bookmarks ===");
        resultBuilder.AppendLine(betweenBookmarks.Trim());

        resultBuilder.AppendLine("\n=== Headers and Footers ===");
        resultBuilder.AppendLine(headerFooterBuilder.ToString().Trim());

        resultBuilder.AppendLine("\n=== Footnotes ===");
        resultBuilder.AppendLine(footnoteBuilder.ToString().Trim());

        // Output the combined result to console.
        Console.WriteLine(resultBuilder.ToString());

        // Optionally, save the extracted content to a text file.
        // Uses the Document.Save rule only for saving documents; here we use standard .NET I/O.
        string outputPath = @"C:\Docs\ExtractedContent.txt";
        File.WriteAllText(outputPath, resultBuilder.ToString());
    }

    // Helper method to safely append header/footer text if the object exists.
    private static void AppendIfNotNull(HeaderFooter hf, StringBuilder sb)
    {
        if (hf != null && hf.HasChildNodes)
        {
            // The Range property gives access to the text of the header/footer.
            sb.AppendLine(hf.Range.Text.Trim());
        }
    }
}
