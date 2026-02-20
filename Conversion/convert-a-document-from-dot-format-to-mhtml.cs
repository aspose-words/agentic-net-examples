using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading; // Added namespace for LoadOptions and LoadFormat

class Program
{
    static void Main()
    {
        // Input DOT (template) file path.
        string inputPath = "Template.dot";

        // Output MHTML file path.
        string outputPath = "Output.mhtml";

        // Load the DOT document. Explicitly specify the load format to ensure correct handling.
        Document doc = new Document(inputPath, new LoadOptions { LoadFormat = LoadFormat.Dot });

        // Configure save options for MHTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use CID URLs for embedded resources (images, fonts, CSS) in the MHTML package.
            ExportCidUrlsForMhtmlResources = true,
            // Choose HTML5 compliance (optional, can be HtmlVersion.Xhtml as well).
            HtmlVersion = HtmlVersion.Html5,
            // Enable pretty formatting for readability (optional).
            PrettyFormat = true
        };

        // Save the document as MHTML.
        doc.Save(outputPath, saveOptions);
    }
}
