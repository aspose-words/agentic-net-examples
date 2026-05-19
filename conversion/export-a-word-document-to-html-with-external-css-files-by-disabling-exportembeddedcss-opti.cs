using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportHtmlWithExternalCss
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello, Aspose.Words!");
        builder.Writeln("This document will be exported to HTML with external CSS.");

        // Save the sample as DOCX to satisfy the bootstrap rule.
        const string inputPath = "sample.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the document from the file we just created.
        Document doc = new Document(inputPath);

        // Configure HtmlSaveOptions to generate an external CSS file.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export CSS to a separate file instead of embedding it.
            CssStyleSheetType = CssStyleSheetType.External
        };

        // Define the output HTML file name.
        const string outputHtml = "output.html";

        // Save the document as HTML using the configured options.
        doc.Save(outputHtml, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputHtml))
            throw new InvalidOperationException("HTML output file was not created.");

        // The external CSS file is saved alongside the HTML file with the same base name.
        string cssPath = Path.ChangeExtension(outputHtml, ".css");
        if (!File.Exists(cssPath))
            throw new InvalidOperationException("External CSS file was not generated.");

        // Display the paths of the generated files.
        Console.WriteLine($"HTML file: {Path.GetFullPath(outputHtml)}");
        Console.WriteLine($"CSS file: {Path.GetFullPath(cssPath)}");
    }
}
