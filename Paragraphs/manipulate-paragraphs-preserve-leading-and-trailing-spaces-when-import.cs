using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class PreserveSpacesExample
{
    static void Main()
    {
        // Sample text with leading and trailing spaces.
        string txtContent = "   First line with spaces   \n" +
                            "  Second line   \n" +
                            "Third line    ";

        // Convert the text to a memory stream.
        using (MemoryStream txtStream = new MemoryStream(Encoding.UTF8.GetBytes(txtContent)))
        {
            // Configure load options to preserve both leading and trailing spaces.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                LeadingSpacesOptions = TxtLeadingSpacesOptions.Preserve,
                TrailingSpacesOptions = TxtTrailingSpacesOptions.Preserve
            };

            // Load the TXT document with the specified options.
            Document txtDoc = new Document(txtStream, loadOptions);

            // Create a new empty DOCX document.
            Document dstDoc = new Document();

            // Append the loaded TXT document into the DOCX document,
            // keeping the source formatting (including spaces).
            dstDoc.AppendDocument(txtDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the resulting DOCX document.
            dstDoc.Save("PreservedSpaces.docx");
        }
    }
}
