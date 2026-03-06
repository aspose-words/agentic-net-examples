using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class TrimSpacesExample
{
    static void Main()
    {
        // Sample text with leading and trailing spaces on each line.
        string textDoc =
            "      Line 1 \n" +
            "    Line 2   \n" +
            " Line 3       ";

        // Configure load options to trim both leading and trailing spaces.
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            LeadingSpacesOptions = TxtLeadingSpacesOptions.Trim,
            TrailingSpacesOptions = TxtTrailingSpacesOptions.Trim
        };

        // Load the TXT content into a Word document using the specified options.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(textDoc)))
        {
            Document doc = new Document(stream, loadOptions);

            // Save the resulting document as DOCX.
            doc.Save("TrimmedSpaces.docx");
        }
    }
}
