using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDotmToMhtml
{
    static void Main()
    {
        // Load the DOTM (macro-enabled template) document.
        Document doc = new Document("input.dotm");

        // Configure save options for MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);

        // Save the document as MHTML.
        doc.Save("output.mhtml", saveOptions);
    }
}
