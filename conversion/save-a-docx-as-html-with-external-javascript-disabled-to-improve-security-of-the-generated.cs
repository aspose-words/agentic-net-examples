using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file.
        const string inputPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX file.
        Document loadedDoc = new Document(inputPath);

        // Configure HTML save options to disable external JavaScript.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            RemoveJavaScriptFromLinks = true
        };

        // Save the document as HTML.
        const string outputPath = "output.html";
        loadedDoc.Save(outputPath, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The HTML output file was not created.");
    }
}
