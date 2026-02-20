// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the DOCX document from file.
        Document doc = new Document("input.docx");

        // StringBuilder to collect extracted text.
        StringBuilder extracted = new StringBuilder();

        // Iterate through all sections to get headers, main body, and footers.
        foreach (Section section in doc.Sections)
        {
            // Headers (all types: primary, first page, even pages).
            foreach (HeaderFooter header in section.HeadersFooters)
            {
                if (header.HeaderFooterType != HeaderFooterType.Footer)
                {
                    extracted.AppendLine(header.GetText());
                }
            }

            // Main body text.
            extracted.AppendLine(section.Body.GetText());

            // Footers (all types).
            foreach (HeaderFooter footer in section.HeadersFooters)
            {
                if (footer.HeaderFooterType == HeaderFooterType.Footer)
                {
                    extracted.AppendLine(footer.GetText());
                }
            }
        }

        // Extract footnotes (including endnotes) text.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnotes)
        {
            extracted.AppendLine(footnote.GetText());
        }

        // Optional: extract text between two bookmarks named "Start" and "End".
        if (doc.Range.Bookmarks["Start"] != null && doc.Range.Bookmarks["End"] != null)
        {
            Bookmark start = doc.Range.Bookmarks["Start"];
            Bookmark end = doc.Range.Bookmarks["End"];

            // Create a range that starts at the bookmark start and ends at the bookmark end.
            Node startNode = start.BookmarkStart;
            Node endNode = end.BookmarkEnd;

            // Build a temporary range.
            Range betweenRange = doc.Range;
            betweenRange.Start = startNode.GetAncestor(NodeType.Body).Document.Range.Start;
            betweenRange.End = endNode.GetAncestor(NodeType.Body).Document.Range.End;

            // Append the text between the bookmarks.
            extracted.AppendLine(betweenRange.GetText());
        }

        // Output the collected text.
        Console.WriteLine(extracted.ToString());
    }
}
