using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Remove the content of all headers and footers in every section.
        // The HeaderFooter objects remain, but they become empty, effectively unlinking them.
        foreach (Section section in doc.Sections)
        {
            section.ClearHeadersFooters();
        }

        // Prepare save options for MHTML format.
        // ExportHeadersFootersMode = None ensures that no header/footer data is written to the output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportHeadersFootersMode = ExportHeadersFootersMode.None
        };

        // Save the document as an MHTML file without headers or footers.
        doc.Save("Output.mht", saveOptions);
    }
}
