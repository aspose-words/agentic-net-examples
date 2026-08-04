using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample content for HTML export.");

        // Save the document locally as DOCX (bootstrap step).
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the saved DOCX file.
        Document doc = new Document(inputPath);

        // Configure HTML save options.
        // Aspose.Words does not expose an ExportJavaScript property; instead,
        // setting RemoveJavaScriptFromLinks disables JavaScript in the output.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            RemoveJavaScriptFromLinks = true
        };

        // Save the document as HTML.
        const string outputPath = "output.html";
        doc.Save(outputPath, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output HTML was not created.");

        Console.WriteLine("Document successfully exported to HTML without JavaScript.");
    }
}
