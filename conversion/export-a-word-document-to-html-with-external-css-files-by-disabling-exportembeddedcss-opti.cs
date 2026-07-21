using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportWordToHtmlWithExternalCss
{
    public static void Main()
    {
        // Create a sample Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");
        builder.Writeln("This document will be exported to HTML with an external CSS file.");

        // Configure HtmlFixedSaveOptions to disable embedded CSS (i.e., use external CSS).
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedCss = false // false => CSS will be saved to an external file.
        };

        // Define output paths.
        string outputHtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "ExportedDocument.html");

        // Save the document as HTML using the configured options.
        doc.Save(outputHtmlPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(outputHtmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // The external CSS file is saved in a folder named after the HTML file (without extension).
        string cssFolder = Path.Combine(Path.GetDirectoryName(outputHtmlPath) ?? string.Empty,
                                        Path.GetFileNameWithoutExtension(outputHtmlPath));
        string cssFilePath = Path.Combine(cssFolder, "styles.css");

        // Validate that the CSS file exists.
        if (!File.Exists(cssFilePath))
            throw new InvalidOperationException("External CSS file was not created.");

        // Optionally, verify that the HTML references the external CSS file.
        string htmlContent = File.ReadAllText(outputHtmlPath);
        if (!htmlContent.Contains("<link rel=\"stylesheet\""))
            throw new InvalidOperationException("HTML does not contain a link to the external CSS file.");

        // Example completed successfully.
        Console.WriteLine("Document exported to HTML with external CSS successfully.");
    }
}
