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

        // Add some content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This is a sample document with fonts embedded as Base64.");

        // Define the output HTML file path.
        string outputPath = "output.html";

        // Configure HtmlSaveOptions to embed fonts as Base64.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportFontsAsBase64 = true,
            CssStyleSheetType = CssStyleSheetType.Embedded,
            PrettyFormat = true
        };

        // Save the document as HTML using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Optionally, you could read the file to ensure it contains Base64 font data.
        // string htmlContent = File.ReadAllText(outputPath);
        // if (!htmlContent.Contains("data:font"))
        //     throw new InvalidOperationException("The HTML does not contain embedded Base64 fonts.");
    }
}
