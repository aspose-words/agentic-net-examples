using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for HTML conversion.");
        builder.Writeln("Another paragraph to generate CSS classes.");

        // Configure HtmlSaveOptions with a CSS class name prefix.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export CSS to an external file so the prefix can be observed.
            CssStyleSheetType = CssStyleSheetType.External,
            // Prefix added to all generated CSS class names.
            CssClassNamePrefix = "myPrefix-",
            // Optional: specify the CSS file name.
            CssStyleSheetFileName = Path.Combine(outputDir, "sample.css")
        };

        // Define the HTML output path.
        string htmlPath = Path.Combine(outputDir, "sample.html");

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // Validate that the external CSS file was created.
        if (!File.Exists(htmlOptions.CssStyleSheetFileName))
            throw new InvalidOperationException("CSS output file was not created.");

        // Optionally, display the paths of the generated files.
        Console.WriteLine($"HTML saved to: {htmlPath}");
        Console.WriteLine($"CSS saved to: {htmlOptions.CssStyleSheetFileName}");
    }
}
