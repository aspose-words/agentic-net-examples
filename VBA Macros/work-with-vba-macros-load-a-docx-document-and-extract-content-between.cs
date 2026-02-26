using System;
using System.Text;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Notes; // Added for Footnote and Endnote types

class ExtractContentExample
{
    static void Main()
    {
        // Load the DOCX document (lifecycle rule: use Document constructor).
        Document doc = new Document("Input.docx");

        // --------------------------------------------------------------------
        // 1. Extract text between two bookmarks (or any two nodes).
        // --------------------------------------------------------------------
        // Assume the document contains bookmarks named "Start" and "End".
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToBookmark("Start");
        Node startNode = builder.CurrentParagraph; // the node where the start bookmark is placed

        builder.MoveToBookmark("End");
        Node endNode = builder.CurrentParagraph; // the node where the end bookmark is placed

        // Collect text from startNode (inclusive) up to endNode (exclusive).
        StringBuilder betweenText = new StringBuilder();
        Node current = startNode;
        while (current != null && current != endNode)
        {
            string nodeText = current.GetText();
            if (!string.IsNullOrEmpty(nodeText))
                betweenText.Append(nodeText);

            current = current.NextPreOrder(doc);
        }

        Console.WriteLine("=== Text Between Bookmarks ===");
        Console.WriteLine(betweenText.ToString());

        // --------------------------------------------------------------------
        // 2. Extract all header and footer text from every section.
        // --------------------------------------------------------------------
        StringBuilder headerFooterText = new StringBuilder();
        foreach (Section section in doc.Sections)
        {
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header != null && !string.IsNullOrWhiteSpace(header.GetText()))
                headerFooterText.AppendLine("Header: " + header.GetText().Trim());

            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer != null && !string.IsNullOrWhiteSpace(footer.GetText()))
                headerFooterText.AppendLine("Footer: " + footer.GetText().Trim());
        }

        Console.WriteLine("\n=== Headers and Footers ===");
        Console.WriteLine(headerFooterText.ToString());

        // --------------------------------------------------------------------
        // 3. Extract all footnote (and endnote) text.
        // --------------------------------------------------------------------
        StringBuilder footnoteText = new StringBuilder();
        NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote fn in footnoteNodes)
        {
            // The Footnote node's GetText() includes the reference mark; trim it for clarity.
            string txt = fn.GetText().Trim();
            if (!string.IsNullOrEmpty(txt))
                footnoteText.AppendLine($"{fn.FootnoteType}: {txt}");
        }

        Console.WriteLine("\n=== Footnotes and Endnotes ===");
        Console.WriteLine(footnoteText.ToString());

        // --------------------------------------------------------------------
        // (Optional) Save the extracted information to a text file.
        // --------------------------------------------------------------------
        File.WriteAllText("ExtractedContent.txt",
            "Between Bookmarks:\r\n" + betweenText + "\r\n\r\n" +
            "Headers and Footers:\r\n" + headerFooterText + "\r\n\r\n" +
            "Footnotes and Endnotes:\r\n" + footnoteText);
    }
}
