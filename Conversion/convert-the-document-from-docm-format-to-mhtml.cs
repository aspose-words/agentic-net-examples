using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the macro-enabled DOCM document.
        Document doc = new Document("input.docm");

        // Configure save options for MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Example: use HTML5 standard for the generated MHTML.
            HtmlVersion = HtmlVersion.Html5,
            // Optional: embed images as Base64 to keep everything in a single file.
            ExportImagesAsBase64 = true,
            // Optional: include document properties in the output.
            ExportDocumentProperties = true
        };

        // Save the document as MHTML.
        doc.Save("output.mhtml", saveOptions);
    }
}
