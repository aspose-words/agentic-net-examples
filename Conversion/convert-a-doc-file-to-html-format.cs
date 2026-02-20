using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputPath = "input.doc";

        // Path to the destination HTML file.
        string outputPath = "output.html";

        // Load the DOC file. The format is auto‑detected from the file extension, so an explicit LoadOptions object is not required.
        Document doc = new Document(inputPath);

        // Set up HTML save options. Use HTML5 and enable pretty formatting.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            HtmlVersion = HtmlVersion.Html5,
            PrettyFormat = true
        };

        // Save the document as HTML.
        doc.Save(outputPath, saveOptions);
    }
}
