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
        builder.Writeln("Hello, Aspose.Words!");
        builder.Writeln("This document will be saved as HTML with external CSS.");

        // Determine output locations.
        string outputHtml = Path.Combine(Directory.GetCurrentDirectory(), "output.html");
        string cssFilePath = Path.ChangeExtension(outputHtml, ".css"); // External CSS file.

        // Configure HtmlSaveOptions to generate an external CSS file.
        // Setting CssStyleSheetType to External creates a separate stylesheet.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External
        };

        // Save the document as HTML with external CSS.
        doc.Save(outputHtml, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputHtml))
            throw new InvalidOperationException("HTML output file was not created.");

        // Verify that the external CSS file was created.
        if (!File.Exists(cssFilePath))
            throw new InvalidOperationException("External CSS file was not created.");

        // Output the locations of the generated files.
        Console.WriteLine($"HTML file saved to: {outputHtml}");
        Console.WriteLine($"External CSS file saved to: {cssFilePath}");
    }
}
