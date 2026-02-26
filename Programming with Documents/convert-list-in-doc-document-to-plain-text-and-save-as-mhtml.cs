using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Update list labels so they reflect the current numbering.
        doc.UpdateListLabels();

        // Set up save options to export list labels as plain text in the MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportListLabels = ExportListLabels.AsInlineText
        };

        // Save the document as MHTML.
        doc.Save("Output.mhtml", saveOptions);
    }
}
