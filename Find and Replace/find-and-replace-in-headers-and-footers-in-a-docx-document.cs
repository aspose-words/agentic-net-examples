using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class HeaderFooterReplace
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

        // Iterate through every section in the document.
        foreach (Section section in doc.Sections)
        {
            // Process all header types.
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.HeaderPrimary], findText, replaceText, options);
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.HeaderFirst],   findText, replaceText, options);
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.HeaderEven],   findText, replaceText, options);

            // Process all footer types.
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.FooterPrimary], findText, replaceText, options);
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.FooterFirst],   findText, replaceText, options);
            ReplaceInHeaderFooter(section.HeadersFooters[HeaderFooterType.FooterEven],   findText, replaceText, options);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Helper method that performs the replace operation on a single HeaderFooter object.
    private static void ReplaceInHeaderFooter(HeaderFooter headerFooter, string pattern, string replacement, FindReplaceOptions options)
    {
        // HeaderFooter may be null if the particular type is not present in the section.
        if (headerFooter == null)
            return;

        // Perform the find‑and‑replace on the header/footer's range.
        headerFooter.Range.Replace(pattern, replacement, options);
    }
}
