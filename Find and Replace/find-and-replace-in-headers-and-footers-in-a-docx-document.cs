using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

class HeaderFooterFindReplace
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Text to find and its replacement.
        string findText = "PLACEHOLDER";
        string replaceText = "Actual Value";

        // Prepare find/replace options (default settings are sufficient for this case).
        FindReplaceOptions options = new FindReplaceOptions();

        // Iterate through all sections in the document.
        foreach (Section section in doc.Sections)
        {
            // Iterate through each header/footer in the current section.
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                // Perform a simple string replace in the header/footer's range.
                // You can also use a regular expression overload if needed.
                headerFooter.Range.Replace(findText, replaceText, options);
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
