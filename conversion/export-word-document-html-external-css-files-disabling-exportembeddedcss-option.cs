using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Configure HTML save options to generate an external CSS file.
        // Setting ExportEmbeddedCss to false disables embedding the CSS into the HTML
        // and causes Aspose.Words to write the stylesheet to a separate .css file.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedCss = false
        };

        // Save the document as HTML. An accompanying .css file will be created
        // in the same folder as the output HTML file.
        doc.Save("Output.html", htmlOptions);
    }
}
