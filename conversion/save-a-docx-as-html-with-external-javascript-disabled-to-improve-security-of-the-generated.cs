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
        builder.Writeln("Sample content for HTML conversion.");
        string docxPath = "sample.docx";
        source.Save(docxPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(docxPath);

        // Configure HTML save options to disable JavaScript in links.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.RemoveJavaScriptFromLinks = true;

        // Save the document as HTML.
        string htmlPath = "output.html";
        doc.Save(htmlPath, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");
    }
}
