using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class HeaderFooterReplaceExample
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        string findText = "_FullName_";
        string replaceText = "John Doe";

        // Configure find/replace options (case‑insensitive, replace whole words).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Iterate through all sections because each section can have its own headers/footers.
        foreach (Section section in doc.Sections)
        {
            // Replace in all header types.
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.HeaderPrimary], findText, replaceText, options);
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.HeaderFirst],   findText, replaceText, options);
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.HeaderEven],   findText, replaceText, options);

            // Replace in all footer types.
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.FooterPrimary], findText, replaceText, options);
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.FooterFirst],   findText, replaceText, options);
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.FooterEven],   findText, replaceText, options);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Helper method that safely performs replace on a HeaderFooter if it exists.
    private static void ReplaceInHeaderFooter(HeaderFooter headerFooter, string find, string replace, FindReplaceOptions options)
    {
        if (headerFooter != null)
        {
            // The Range object of the header/footer contains its text.
            headerFooter.Range.Replace(find, replace, options);
        }
    }
}
