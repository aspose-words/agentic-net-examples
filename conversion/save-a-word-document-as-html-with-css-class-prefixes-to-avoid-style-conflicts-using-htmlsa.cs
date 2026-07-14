using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output file names.
        string outputHtml = "output.html";
        string outputCss = Path.ChangeExtension(outputHtml, ".css");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading styled paragraph.
        // Use the built‑in "Header" style so that the generated CSS class will be "Header".
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Header;
        builder.Writeln("Sample Heading");

        // Add a normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a sample paragraph.");

        // Configure HTML save options with an external CSS file and a class name prefix.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External,
            CssClassNamePrefix = "myPrefix-"
        };

        // Save the document as HTML.
        doc.Save(outputHtml, saveOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(outputHtml))
            throw new InvalidOperationException($"HTML file '{outputHtml}' was not created.");

        // Validate that the external CSS file was created.
        if (!File.Exists(outputCss))
            throw new InvalidOperationException($"CSS file '{outputCss}' was not created.");

        // Simple content checks to ensure the prefix was applied.
        string htmlContent = File.ReadAllText(outputHtml);
        if (!htmlContent.Contains("class=\"myPrefix-Header\""))
            throw new InvalidOperationException("CSS class prefix was not applied to the heading in the HTML output.");

        string cssContent = File.ReadAllText(outputCss);
        if (!cssContent.Contains(".myPrefix-Header"))
            throw new InvalidOperationException("CSS class prefix was not applied in the CSS file.");

        // Indicate successful completion.
        Console.WriteLine("Document saved as HTML with CSS class prefix successfully.");
    }
}
