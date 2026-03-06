using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOCX template from disk.
            Document doc = new Document("Template.docx");

            // Configure the save options for the fixed‑page HTML format.
            HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
            {
                // The format is automatically set to HtmlFixed by the options class.
                // Additional options can be set here if needed, e.g.:
                // ExportEmbeddedImages = false,
                // ResourcesFolder = "Resources"
            };

            // Save the document as fixed‑page HTML.
            doc.Save("Output.html", htmlOptions);
        }
    }
}
