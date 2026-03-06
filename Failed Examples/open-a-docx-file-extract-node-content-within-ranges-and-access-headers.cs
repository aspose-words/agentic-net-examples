// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file. The constructor automatically detects the format.
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // 1. Extract text from the whole document.
        // -----------------------------------------------------------------
        string wholeDocumentText = doc.GetText(); // Returns all text including paragraph breaks.
        Console.WriteLine("=== Whole Document Text ===");
        Console.WriteLine(wholeDocumentText);

        // -----------------------------------------------------------------
        // 2. Extract text from a specific range – here the first paragraph.
        // -----------------------------------------------------------------
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        string firstParagraphText = firstParagraph.GetText(); // Includes the paragraph break.
        Console.WriteLine("\n=== First Paragraph Text ===");
        Console.WriteLine(firstParagraphText);

        // -----------------------------------------------------------------
        // 3. Access header and footer contents (primary type).
        // -----------------------------------------------------------------
        HeaderFooter header = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (header != null)
        {
            Console.WriteLine("\n=== Header Text (Primary) ===");
            Console.WriteLine(header.GetText());
        }

        HeaderFooter footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
        if (footer != null)
        {
            Console.WriteLine("\n=== Footer Text (Primary) ===");
            Console.WriteLine(footer.GetText());
        }

        // -----------------------------------------------------------------
        // 4. Access footnotes (and endnotes) in the document.
        // -----------------------------------------------------------------
        NodeCollection footnoteNodes = doc.GetChildNodes(NodeType.Footnote, true);
        Console.WriteLine($"\n=== Footnotes Found: {footnoteNodes.Count} ===");
        foreach (Footnote footnote in footnoteNodes)
        {
            // Footnote.GetText() returns the footnote text plus the footnote mark.
            string footnoteText = footnote.GetText().Trim();
            Console.WriteLine($"[{footnote.FootnoteType}] {footnoteText}");
        }

        // -----------------------------------------------------------------
        // 5. Save the document (optional – demonstrates the required save rule).
        // -----------------------------------------------------------------
        doc.Save("Output.docx"); // Save format inferred from the file extension.
    }
}
