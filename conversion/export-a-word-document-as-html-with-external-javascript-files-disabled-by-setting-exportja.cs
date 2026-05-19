using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document exported to HTML without JavaScript.");

        // Configure HTML save options to prevent JavaScript from being exported.
        // The ExportJavaScript property is not available in recent versions; instead,
        // we can remove JavaScript from links, which ensures no script references are written.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            RemoveJavaScriptFromLinks = true
        };

        // Save the document as HTML.
        string outputPath = "output.html";
        doc.Save(outputPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // Ensure that no JavaScript was exported.
        string htmlContent = File.ReadAllText(outputPath);
        if (htmlContent.Contains("<script", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("JavaScript was not disabled in the exported HTML.");
    }
}
