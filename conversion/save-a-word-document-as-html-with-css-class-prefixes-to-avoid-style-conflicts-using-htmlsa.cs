using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");
        builder.Writeln("This document will be saved as HTML with a CSS class prefix.");

        // Prepare the HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Export CSS to an external stylesheet so we can see the prefix in the .css file.
            CssStyleSheetType = CssStyleSheetType.External,
            // Add a prefix to every generated CSS class name.
            CssClassNamePrefix = "myPrefix-"
        };

        // Define output file names.
        string htmlPath = "output.html";
        string cssPath = "output.css";

        // Ensure any previous files are removed.
        if (File.Exists(htmlPath)) File.Delete(htmlPath);
        if (File.Exists(cssPath)) File.Delete(cssPath);

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // When CssStyleSheetType is External, Aspose.Words creates a CSS file with the same name as the HTML file.
        // The default name is the HTML file name with a .css extension.
        if (!File.Exists(cssPath))
            throw new InvalidOperationException("CSS output file was not created.");

        // Optional: read the generated files to ensure the prefix appears.
        string htmlContent = File.ReadAllText(htmlPath);
        string cssContent = File.ReadAllText(cssPath);

        if (!cssContent.Contains(".myPrefix-"))
            throw new InvalidOperationException("CSS class prefix was not applied.");

        // Indicate success (no console input required).
        Console.WriteLine("HTML and CSS files were generated successfully.");
    }
}
