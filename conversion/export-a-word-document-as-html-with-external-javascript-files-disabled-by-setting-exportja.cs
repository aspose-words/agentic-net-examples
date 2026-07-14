using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document.");
        // Insert a hyperlink that contains JavaScript.
        builder.InsertHyperlink("Click me", "javascript:alert('Hello')", false);

        // Define the output HTML file path.
        string outputPath = "output.html";

        // Configure HTML save options to remove JavaScript from links.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // This option strips JavaScript from hyperlink URLs.
            RemoveJavaScriptFromLinks = true
        };

        // Save the document as HTML.
        doc.Save(outputPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("HTML output file was not created.");
    }
}
