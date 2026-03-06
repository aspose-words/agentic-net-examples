using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class TxtToDocxExample
{
    static void Main()
    {
        // Sample text that contains leading and trailing spaces on each line.
        string txt = "   First line with leading spaces   \n" +
                     "\tSecond line with a tab at start and spaces at end   \n" +
                     "Third line without extra spaces";

        // Convert the string to a UTF‑8 byte array and place it in a MemoryStream.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(txt)))
        {
            // Create TxtLoadOptions and set the desired handling for leading and trailing spaces.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                // Preserve all leading spaces (do not convert to indent or trim).
                LeadingSpacesOptions = TxtLeadingSpacesOptions.Preserve,
                // Preserve all trailing spaces (do not trim).
                TrailingSpacesOptions = TxtTrailingSpacesOptions.Preserve
            };

            // Load the TXT content into a Document using the stream and the configured options.
            Document doc = new Document(stream, loadOptions);

            // Save the resulting document as a DOCX file.
            doc.Save("Output.docx");
        }
    }
}
