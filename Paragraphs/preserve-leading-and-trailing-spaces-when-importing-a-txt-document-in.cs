using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class PreserveSpacesExample
{
    static void Main()
    {
        // Sample text with leading and trailing spaces on each line.
        string txtContent = "    First line with leading spaces and trailing spaces   \n" +
                            "  Second line  \n" +
                            "Third line without extra spaces";

        // Convert the string to a UTF-8 encoded memory stream.
        using (MemoryStream txtStream = new MemoryStream(Encoding.UTF8.GetBytes(txtContent)))
        {
            // Configure load options to preserve both leading and trailing spaces.
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                LeadingSpacesOptions = TxtLeadingSpacesOptions.Preserve,
                TrailingSpacesOptions = TxtTrailingSpacesOptions.Preserve
            };

            // Load the TXT document into an Aspose.Words Document using the specified options.
            Document doc = new Document(txtStream, loadOptions);

            // Save the resulting document as DOCX, preserving the spaces.
            doc.Save("PreservedSpaces.docx", SaveFormat.Docx);
        }
    }
}
