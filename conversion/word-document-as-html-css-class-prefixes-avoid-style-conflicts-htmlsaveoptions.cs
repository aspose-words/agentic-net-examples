using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Configure HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Export CSS to an external stylesheet file.
            CssStyleSheetType = CssStyleSheetType.External,

            // Add a prefix to every generated CSS class name to prevent naming conflicts.
            CssClassNamePrefix = "myapp-",

            // Specify the name (and optionally the path) of the external CSS file.
            CssStyleSheetFileName = "Output.css"
        };

        // Save the document as HTML using the configured options.
        doc.Save("Output.html", saveOptions);
    }
}
