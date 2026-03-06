using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Notes; // Added for Footnote class

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // List to hold extracted text fragments.
        List<string> extractedTexts = new List<string>();

        // 1. Extract text from the main document body (the whole range).
        extractedTexts.Add(doc.Range.Text.Trim());

        // 2. Extract text from all headers and footers in each section.
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                // HeaderFooter.Range gives the range of the header/footer.
                string headerFooterText = headerFooter.Range.Text.Trim();
                if (!string.IsNullOrEmpty(headerFooterText))
                {
                    extractedTexts.Add(headerFooterText);
                }
            }
        }

        // 3. Extract text from all footnotes (and endnotes) in the document.
        NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);
        foreach (Footnote footnote in footnotes)
        {
            // Footnote.Range contains the content of the footnote.
            string footnoteText = footnote.Range.Text.Trim();
            if (!string.IsNullOrEmpty(footnoteText))
            {
                extractedTexts.Add(footnoteText);
            }
        }

        // (Optional) Output the collected texts to the console.
        Console.WriteLine("Extracted Text Fragments:");
        foreach (string text in extractedTexts)
        {
            Console.WriteLine("---");
            Console.WriteLine(text);
        }

        // Save the (unchanged) document to a new file if needed.
        doc.Save("output.docx");
    }
}
