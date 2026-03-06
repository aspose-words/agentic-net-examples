using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOC document that contains the list.
        Document doc = new Document("Input.doc");

        // Update list labels so that they are correct before exporting.
        doc.UpdateListLabels();

        // Configure HTML saving to render list labels as plain‑text (inline) rather than HTML list tags.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportListLabels = ExportListLabels.AsInlineText
        };

        // Save the document as HTML. The resulting HTML will contain the list items as plain text.
        doc.Save("Output.html", htmlOptions);
    }
}
