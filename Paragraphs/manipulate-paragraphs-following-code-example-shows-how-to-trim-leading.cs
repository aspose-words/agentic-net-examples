using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Sample text containing leading and trailing spaces on each line.
        string text = "   Line 1   \n" +
                      "\tLine 2\t\n" +
                      "  Line 3  ";

        // Convert the string to a UTF‑8 byte array and wrap it in a memory stream.
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        using (MemoryStream stream = new MemoryStream(bytes))
        {
            // Set load options to trim both leading and trailing spaces.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                LeadingSpacesOptions = TxtLeadingSpacesOptions.Trim,
                TrailingSpacesOptions = TxtTrailingSpacesOptions.Trim
            };

            // Load the TXT content into a Word document using the specified options.
            Document doc = new Document(stream, loadOptions);

            // Save the resulting document as DOCX.
            doc.Save("TrimmedOutput.docx");
        }
    }
}
