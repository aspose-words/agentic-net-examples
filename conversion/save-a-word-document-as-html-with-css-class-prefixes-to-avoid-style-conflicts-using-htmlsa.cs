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
        builder.Writeln("Hello World!");
        builder.Writeln("This is a sample paragraph.");

        // Configure HTML save options:
        // - Export CSS to an external file.
        // - Add a prefix to all generated CSS class names to avoid conflicts.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External,
            CssClassNamePrefix = "myPrefix-"
        };

        // Prepare output paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string htmlPath = Path.Combine(outputDir, "Sample.html");
        string expectedCssPath = Path.ChangeExtension(htmlPath, ".css");

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, saveOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Validate that the external CSS file was created.
        if (!File.Exists(expectedCssPath))
            throw new InvalidOperationException("The CSS output file was not created.");

        // Verify that the CSS class prefix appears in the HTML content.
        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("class=\"myPrefix-"))
            throw new InvalidOperationException("The CSS class prefix was not applied to the HTML.");

        // Example completed successfully.
    }
}
