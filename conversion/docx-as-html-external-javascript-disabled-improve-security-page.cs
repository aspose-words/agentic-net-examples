using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a simple paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world! This is a sample DOCX converted to HTML.");

        // Configure HTML save options.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Disable JavaScript in links for security.
            RemoveJavaScriptFromLinks = true
        };

        // Save the document as HTML using the configured options.
        doc.Save("OutputDocument.html", htmlOptions);

        Console.WriteLine("Document successfully saved as OutputDocument.html");
    }
}
