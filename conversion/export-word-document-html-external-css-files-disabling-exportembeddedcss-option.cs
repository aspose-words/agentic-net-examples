using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Always create a new document to avoid issues with existing files that may contain unsupported fields.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Configure HTML save options to generate an external CSS file.
        var htmlOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External
        };

        // Save the document as HTML. An accompanying .css file will be created
        // in the same folder as the output HTML file.
        doc.Save("Output.html", htmlOptions);
    }
}
