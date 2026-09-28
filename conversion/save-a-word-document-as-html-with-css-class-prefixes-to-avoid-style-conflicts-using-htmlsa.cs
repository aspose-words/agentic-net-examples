using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document.");

        // Define the output HTML file path.
        string outputPath = "output.html";

        // Configure HTML save options to add a CSS class name prefix.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.CssClassNamePrefix = "myPrefix_";

        // Save the document as HTML using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("HTML output file was not created.");
        }
    }
}
