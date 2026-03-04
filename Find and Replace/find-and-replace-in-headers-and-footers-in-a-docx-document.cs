using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class HeaderFooterReplace
{
    static void Main()
    {
        // Input and output file paths.
        string inputPath = @"C:\Docs\InputDocument.docx";
        string outputPath = @"C:\Docs\OutputDocument.docx";

        // Text to find and its replacement.
        string findText = "_FullName_";
        string replaceText = "John Doe";

        // Load the document (uses the provided load rule).
        Document doc = new Document(inputPath);

        // Configure find/replace options (case‑insensitive, replace whole words).
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Iterate through every section and replace text in all header/footer types.
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooterType hfType in Enum.GetValues(typeof(HeaderFooterType)))
            {
                HeaderFooter headerFooter = section.HeadersFooters[hfType];
                if (headerFooter != null)
                {
                    // Perform the replace operation on the header/footer's range.
                    headerFooter.Range.Replace(findText, replaceText, options);
                }
            }
        }

        // Save the modified document (uses the provided save rule).
        doc.Save(outputPath);
    }
}
