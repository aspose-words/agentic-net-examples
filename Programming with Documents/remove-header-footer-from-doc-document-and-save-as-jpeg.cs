using Aspose.Words;
using Aspose.Words.Saving;
using System;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Remove all headers and footers from every section.
        foreach (Section section in doc.Sections)
        {
            section.HeadersFooters.Clear();
        }

        // Set up image save options for JPEG format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        saveOptions.JpegQuality = 90; // Adjust quality as needed (0‑100).

        // Save the first page of the document as a JPEG image.
        doc.Save("Output.jpg", saveOptions);
    }
}
