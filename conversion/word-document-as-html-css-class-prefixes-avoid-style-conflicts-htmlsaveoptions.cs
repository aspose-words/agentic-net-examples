using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create HtmlSaveOptions for HTML output.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Export CSS to an external stylesheet file.
        saveOptions.CssStyleSheetType = CssStyleSheetType.External;

        // Add a prefix to every generated CSS class name to prevent naming conflicts.
        saveOptions.CssClassNamePrefix = "myapp-";

        // Specify the name (and optionally the path) of the external CSS file.
        saveOptions.CssStyleSheetFileName = "Output.css";

        // Save the document as HTML using the configured options.
        doc.Save("Output.html", saveOptions);
    }
}
