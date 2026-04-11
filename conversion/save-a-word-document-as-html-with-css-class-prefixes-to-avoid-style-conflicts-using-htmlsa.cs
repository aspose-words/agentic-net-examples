using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create an output directory relative to the current working directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Heading");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a sample paragraph to demonstrate CSS class prefix.");

        // Configure HTML save options:
        // - Export CSS to an external file.
        // - Add a prefix to all generated CSS class names.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External,
            CssClassNamePrefix = "myprefix-"
        };

        // Define the output HTML file path.
        string htmlPath = Path.Combine(outputDir, "Sample.html");

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, saveOptions);

        // The external CSS file will be created alongside the HTML file with the same base name.
        string cssPath = Path.ChangeExtension(htmlPath, ".css");

        // Validate that both files were created.
        if (!File.Exists(htmlPath))
            throw new FileNotFoundException("HTML output file was not created.", htmlPath);
        if (!File.Exists(cssPath))
            throw new FileNotFoundException("CSS output file was not created.", cssPath);

        // Verify that the CSS file contains the expected class name prefix.
        string cssContent = File.ReadAllText(cssPath);
        if (!cssContent.Contains(".myprefix-"))
            throw new InvalidOperationException("CSS class name prefix was not applied.");

        // Indicate successful execution.
        Console.WriteLine("HTML and CSS files generated successfully.");
    }
}
