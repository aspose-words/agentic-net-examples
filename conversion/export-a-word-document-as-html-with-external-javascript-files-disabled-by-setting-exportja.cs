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

        // Configure HTML save options.
        // The ExportJavaScript property is not available in the current Aspose.Words version.
        // By default, Aspose.Words does not generate external JavaScript files, so we simply use the default options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Save the document as HTML.
        string outputPath = "output.html";
        doc.Save(outputPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output HTML was not created.");
        }

        Console.WriteLine($"HTML file successfully created at: {Path.GetFullPath(outputPath)}");
    }
}
