using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractImagesToBase64Html
{
    static void Main()
    {
        // Path to the source DOCX file.
        const string sourceDocx = @"C:\Docs\SourceDocument.docx";

        // Load the Word document from the file system.
        Document doc = new Document(sourceDocx);

        // Configure HTML save options to embed images as Base64 data.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportImagesAsBase64 = true,
            PrettyFormat = true
        };

        // Save the document to a memory stream using the configured options.
        using (MemoryStream htmlStream = new MemoryStream())
        {
            doc.Save(htmlStream, htmlOptions);

            // Convert the stream contents to a UTF‑8 string containing the HTML.
            string htmlContent = Encoding.UTF8.GetString(htmlStream.ToArray());

            // Write the resulting HTML to a file (or use it as needed).
            const string outputHtml = @"C:\Docs\DocumentWithEmbeddedImages.html";
            File.WriteAllText(outputHtml, htmlContent, Encoding.UTF8);
        }
    }
}
