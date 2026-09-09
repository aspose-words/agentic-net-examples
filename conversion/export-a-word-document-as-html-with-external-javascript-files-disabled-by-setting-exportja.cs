using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add a line of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words HTML export with JavaScript disabled.");

        // Configure HTML save options. The ExportJavaScript property no longer exists,
        // so we simply use the default options which do not embed JavaScript.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Define the output HTML file name.
        string outputFile = "output.html";

        // Save the document as HTML using the configured options.
        doc.Save(outputFile, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Expected HTML output file was not created.");
    }
}
