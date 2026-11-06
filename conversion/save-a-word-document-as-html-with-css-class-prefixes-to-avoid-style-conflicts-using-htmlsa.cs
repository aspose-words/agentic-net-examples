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
        builder.Writeln("It will be saved as HTML with CSS class prefixes.");

        // Apply formatting that will cause Aspose.Words to generate a CSS class.
        builder.Font.Bold = true;
        builder.Writeln("This line is bold and will trigger a CSS class.");

        // Configure HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        // Use an external (or embedded) stylesheet so that CSS classes are generated.
        saveOptions.CssStyleSheetType = CssStyleSheetType.External;
        // Prefix for generated CSS class names.
        saveOptions.CssClassNamePrefix = "myPrefix_";

        // Save the document as HTML.
        string outputPath = "output.html";
        doc.Save(outputPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output HTML was not created.");

        // Ensure the CSS class prefix appears in the generated HTML.
        string htmlContent = File.ReadAllText(outputPath);
        if (!htmlContent.Contains("myPrefix_"))
            throw new InvalidOperationException("CSS class name prefix was not applied in the HTML output.");
    }
}
