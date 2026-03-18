using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class HeaderKeywordReplace
{
    static void Main()
    {
        // Paths to the source and destination documents.
        string inputPath = @"C:\Docs\Input.docx";
        string outputPath = @"C:\Docs\Output.docx";

        // The keyword that must be present in a header for the replacement to occur.
        const string headerKeyword = "CONFIDENTIAL";

        // Text to find and its replacement.
        const string textToFind = "OldCompany";
        const string textToReplace = "NewCompany";

        // Load the document (lifecycle rule: load).
        Document doc = new Document(inputPath);

        // Iterate through all sections and their headers.
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                // Process only headers (not footers) that contain the specific keyword.
                if (headerFooter.IsHeader && headerFooter.GetText().Contains(headerKeyword))
                {
                    // Configure replace options if needed (e.g., case‑insensitive).
                    FindReplaceOptions options = new FindReplaceOptions
                    {
                        MatchCase = false,
                        FindWholeWordsOnly = false
                    };

                    // Perform the replacement within this header's range.
                    headerFooter.Range.Replace(textToFind, textToReplace, options);
                }
            }
        }

        // Save the modified document (lifecycle rule: save).
        doc.Save(outputPath);
    }
}
