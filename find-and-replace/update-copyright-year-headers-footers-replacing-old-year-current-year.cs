using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class UpdateHeaderFooterYear
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("Input.docx");

        // Current year to replace old years with.
        string currentYear = DateTime.Now.Year.ToString();

        // Options for the find‑replace operation.
        FindReplaceOptions options = new FindReplaceOptions
        {
            MatchCase = false,
            FindWholeWordsOnly = false
        };

        // Iterate through every section and every header/footer within it.
        foreach (Section section in doc.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                // Replace any four‑digit number (a year) with the current year.
                headerFooter.Range.Replace(new Regex(@"\b\d{4}\b"), currentYear, options);
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
