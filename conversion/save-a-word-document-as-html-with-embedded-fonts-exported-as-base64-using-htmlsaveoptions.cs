using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text using a specific font to ensure font data is present.
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample paragraph with the Arial font.");

        // Configure HTML save options to embed fonts as Base64.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportFontsAsBase64 = true,
            CssStyleSheetType = CssStyleSheetType.Embedded,
            PrettyFormat = true
        };

        // Define the output HTML file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.html");

        // Save the document as HTML with the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Optionally, indicate success (no interactive prompts required).
        Console.WriteLine("Document saved successfully to: " + outputPath);
    }
}
