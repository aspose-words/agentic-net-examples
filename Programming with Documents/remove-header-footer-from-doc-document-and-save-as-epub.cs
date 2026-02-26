using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Prepare save options for EPUB format and disable header/footer export.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub);
        saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.None;
        saveOptions.Encoding = Encoding.UTF8; // optional, ensures UTF‑8 output.

        // Save the document as EPUB without headers and footers.
        doc.Save("Output.epub", saveOptions);
    }
}
