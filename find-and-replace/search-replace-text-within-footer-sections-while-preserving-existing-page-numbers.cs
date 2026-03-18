using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FooterReplaceExample
{
    static void Main()
    {
        // Load the existing document.
        Document doc = new Document("Input.docx");

        // Iterate through all sections because each section can have its own footers.
        foreach (Section section in doc.Sections)
        {
            // Access the collection of headers and footers for the current section.
            HeaderFooterCollection footers = section.HeadersFooters;

            // Process the three possible footer types (Primary, First page, Even pages).
            HeaderFooterType[] types = {
                HeaderFooterType.FooterPrimary,
                HeaderFooterType.FooterFirst,
                HeaderFooterType.FooterEven
            };

            foreach (HeaderFooterType type in types)
            {
                HeaderFooter footer = footers[type];
                if (footer == null) continue; // No footer of this type in the current section.

                // Configure find/replace options.
                // - IgnoreFields = true ensures that fields such as PAGE (page numbers) are not altered.
                // - MatchCase and FindWholeWordsOnly can be set according to requirements.
                FindReplaceOptions options = new FindReplaceOptions
                {
                    MatchCase = false,
                    FindWholeWordsOnly = false,
                    IgnoreFields = true
                };

                // Perform the replacement within the footer's range.
                // Replace "PLACEHOLDER" with the desired text while leaving page numbers untouched.
                footer.Range.Replace("PLACEHOLDER", "Actual Text", options);
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
