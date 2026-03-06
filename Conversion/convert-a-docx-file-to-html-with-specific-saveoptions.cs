using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("input.docx");

        // Create HtmlSaveOptions and configure desired settings.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Export images as Base64 strings embedded in the HTML.
            ExportImagesAsBase64 = true,

            // Produce nicely indented (pretty) HTML output.
            PrettyFormat = true,

            // Use UTF-8 encoding without a BOM.
            Encoding = new UTF8Encoding(false)
        };

        // Save the document as HTML using the configured options.
        doc.Save("output.html", htmlOptions);
    }
}
