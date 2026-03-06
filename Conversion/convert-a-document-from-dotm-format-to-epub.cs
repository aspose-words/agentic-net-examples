using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOTM file.
        Document doc = new Document("InputTemplate.dotm");

        // Prepare EPUB save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.SaveFormat = SaveFormat.Epub;   // Specify EPUB output.
        saveOptions.Encoding = Encoding.UTF8;       // Use UTF‑8 encoding.
        saveOptions.ExportDocumentProperties = true; // Optional: include document properties.

        // Save the document as EPUB.
        doc.Save("OutputDocument.epub", saveOptions);
    }
}
