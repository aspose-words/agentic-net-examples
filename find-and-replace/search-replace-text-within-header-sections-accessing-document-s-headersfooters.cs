using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class HeaderReplaceExample
{
    static void Main()
    {
        // Load an existing document (replace with your file path)
        Document doc = new Document(@"C:\Docs\Input.docx");

        // Text to find and its replacement
        string findText = "_FullName_";
        string replaceText = "John Doe";

        // Options for the replace operation (case‑insensitive, replace whole words)
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Iterate through every section in the document
        foreach (Section section in doc.Sections)
        {
            // Access the collection of headers/footers for the current section
            HeaderFooterCollection headersFooters = section.HeadersFooters;

            // Process each type of header that may exist
            ReplaceInHeader(headersFooters, HeaderFooterType.HeaderPrimary, findText, replaceText, options);
            ReplaceInHeader(headersFooters, HeaderFooterType.HeaderFirst,   findText, replaceText, options);
            ReplaceInHeader(headersFooters, HeaderFooterType.HeaderEven,   findText, replaceText, options);
        }

        // Save the modified document (replace with your desired output path)
        doc.Save(@"C:\Docs\Output.docx");
    }

    // Helper method that performs the replace on a specific header type if it exists
    private static void ReplaceInHeader(HeaderFooterCollection collection,
                                         HeaderFooterType headerType,
                                         string pattern,
                                         string replacement,
                                         FindReplaceOptions options)
    {
        // Retrieve the header of the specified type; may be null if not present
        HeaderFooter header = collection[headerType];
        if (header != null)
        {
            // Perform the find‑and‑replace on the header's range
            header.Range.Replace(pattern, replacement, options);
        }
    }
}
