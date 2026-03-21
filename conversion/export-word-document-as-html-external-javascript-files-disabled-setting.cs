using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Always create a simple document to avoid issues with unknown input files.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document created for the HTML export example.");

        // Configure HTML save options to remove any JavaScript from links.
        var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            RemoveJavaScriptFromLinks = true
        };

        // Save the document as HTML.
        const string outputPath = "output.html";
        doc.Save(outputPath, htmlOptions);

        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
