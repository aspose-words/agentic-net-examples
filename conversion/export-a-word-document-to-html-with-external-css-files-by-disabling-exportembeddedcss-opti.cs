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
        builder.Font.Size = 24;
        builder.Font.Bold = true;
        builder.Writeln("Sample Heading");
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.Writeln("This is a paragraph with some text.");

        // Define output file names.
        string htmlPath = "output.html";
        string cssPath = "output.css";

        // Configure HTML save options to generate an external CSS file.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            CssStyleSheetFileName = cssPath,
            // Ensure the stylesheet is written externally.
            CssStyleSheetType = CssStyleSheetType.External
        };

        // Save the document as HTML with external CSS.
        doc.Save(htmlPath, saveOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException($"HTML file '{htmlPath}' was not created.");

        // Validate that the CSS file was created.
        if (!File.Exists(cssPath))
            throw new InvalidOperationException($"CSS file '{cssPath}' was not created.");

        // Ensure both files have content.
        FileInfo htmlInfo = new FileInfo(htmlPath);
        FileInfo cssInfo = new FileInfo(cssPath);

        if (htmlInfo.Length == 0)
            throw new InvalidOperationException("HTML file is empty.");

        if (cssInfo.Length == 0)
            throw new InvalidOperationException("CSS file is empty.");

        Console.WriteLine("HTML and external CSS files were successfully created.");
    }
}
