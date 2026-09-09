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

        // Configure HTML Fixed save options to disable embedded CSS,
        // which causes Aspose.Words to generate an external CSS file.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedCss = false
        };

        // Define output paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        string htmlPath = Path.Combine(outputDir, "sample.html");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Save the document as HTML with the specified options.
        doc.Save(htmlPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // The external CSS is saved in a folder named after the HTML file (without extension).
        string cssFolder = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(htmlPath));
        string cssPath = Path.Combine(cssFolder, "styles.css");

        // Validate that the external CSS file exists.
        if (!File.Exists(cssPath))
            throw new InvalidOperationException("External CSS file was not created.");

        // Optional: verify that the HTML references the external CSS file.
        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("<link") || !htmlContent.Contains("styles.css"))
            throw new InvalidOperationException("HTML does not reference the external CSS file.");

        // Indicate successful conversion.
        Console.WriteLine("Document successfully exported to HTML with external CSS.");
    }
}
