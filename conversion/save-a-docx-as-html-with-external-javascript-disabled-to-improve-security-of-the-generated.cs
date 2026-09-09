using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample content for conversion to HTML.");
        source.Save("input.docx", SaveFormat.Docx);

        // Load the created DOCX.
        Document doc = new Document("input.docx");

        // Configure HTML save options to remove JavaScript from links.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            RemoveJavaScriptFromLinks = true
        };

        // Save the document as HTML.
        string htmlPath = "output.html";
        doc.Save(htmlPath, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Optionally, you could read the file to ensure it contains expected content.
        // string htmlContent = File.ReadAllText(htmlPath);
        // Console.WriteLine(htmlContent);
    }
}
