// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple bullet list in the main body of the document.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Move the builder to the end of the document body.
        builder.MoveToDocumentEnd();

        // Start a bullet list.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");
        // End the list.
        builder.ListFormat.RemoveNumbers();

        // -----------------------------------------------------------------
        // 2. Extract text between two specific nodes.
        //    For demonstration, we locate two paragraphs by their text.
        // -----------------------------------------------------------------
        // Find the start paragraph.
        Paragraph startParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        // Find the end paragraph (the last paragraph in the document).
        Paragraph endParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, doc.GetChildNodes(NodeType.Paragraph, true).Count - 1, true);

        // Collect text of all nodes that lie between startParagraph (exclusive) and endParagraph (exclusive).
        List<string> betweenTexts = new List<string>();
        Node current = startParagraph.NextSibling;
        while (current != null && current != endParagraph)
        {
            // Use the Range property to get the text of the current node.
            betweenTexts.Add(current.Range.Text);
            current = current.NextSibling;
        }

        // Output the extracted text to the console.
        Console.WriteLine("Text between the selected nodes:");
        foreach (string txt in betweenTexts)
        {
            Console.WriteLine(txt.Trim());
        }

        // -----------------------------------------------------------------
        // 3. Process headers and footers of each section.
        // -----------------------------------------------------------------
        Console.WriteLine("\nHeaders and Footers:");
        foreach (Section section in doc.Sections)
        {
            // Header (primary)
            HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header != null && !string.IsNullOrWhiteSpace(header.GetText()))
                Console.WriteLine("Header: " + header.GetText().Trim());

            // Footer (primary)
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer != null && !string.IsNullOrWhiteSpace(footer.GetText()))
                Console.WriteLine("Footer: " + footer.GetText().Trim());
        }

        // -----------------------------------------------------------------
        // 4. Process footnotes in the document.
        // -----------------------------------------------------------------
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        Console.WriteLine("\nFootnotes:");
        foreach (Footnote footnote in footnotes)
        {
            // The Footnote node's Range contains its text.
            Console.WriteLine($"Footnote ({footnote.FootnoteType}): {footnote.Range.Text.Trim()}");
        }

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        doc.Save("Output.docx");
    }
}
