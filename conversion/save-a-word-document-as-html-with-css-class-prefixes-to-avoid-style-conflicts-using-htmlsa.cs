using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HtmlConversionWithCssPrefix
{
    public static void Main()
    {
        // Define file names.
        const string docFileName = "sample.docx";
        const string htmlFileName = "sample.html";

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");
        doc.Save(docFileName, SaveFormat.Docx);

        // Load the document we just created.
        Document loadedDoc = new Document(docFileName);

        // Configure HTML save options with a CSS class name prefix.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External, // Export CSS to an external file.
            CssClassNamePrefix = "myPrefix-"               // Prefix added to all generated CSS classes.
        };

        // Save the document as HTML using the configured options.
        loadedDoc.Save(htmlFileName, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlFileName))
            throw new InvalidOperationException($"HTML file '{htmlFileName}' was not created.");

        // Verify that the external CSS file was also created.
        string cssFileName = Path.ChangeExtension(htmlFileName, ".css");
        if (!File.Exists(cssFileName))
            throw new InvalidOperationException($"CSS file '{cssFileName}' was not created.");

        // Optional: output the locations of the generated files.
        Console.WriteLine($"HTML saved to: {Path.GetFullPath(htmlFileName)}");
        Console.WriteLine($"CSS saved to: {Path.GetFullPath(cssFileName)}");
    }
}
