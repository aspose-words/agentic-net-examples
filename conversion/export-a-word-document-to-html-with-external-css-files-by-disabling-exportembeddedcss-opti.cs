using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        string docPath = Path.Combine(outputDir, "Sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Load the document back (bootstrap rule for existing DOCX).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // 3. Configure HtmlFixedSaveOptions to disable embedded CSS.
        //    This forces Aspose.Words to write an external CSS file.
        // -----------------------------------------------------------------
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedCss = false // false => external CSS file.
        };

        // -----------------------------------------------------------------
        // 4. Save the document as HTML.
        // -----------------------------------------------------------------
        string htmlPath = Path.Combine(outputDir, "Sample.html");
        loadedDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the HTML file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new FileNotFoundException("HTML output file was not created.", htmlPath);

        // -----------------------------------------------------------------
        // 6. Validate that the external CSS file was created.
        //    Aspose.Words creates a folder named after the HTML file (without extension)
        //    and places a file named "styles.css" inside it.
        // -----------------------------------------------------------------
        string cssFolder = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(htmlPath));
        string cssPath = Path.Combine(cssFolder, "styles.css");

        if (!File.Exists(cssPath))
            throw new FileNotFoundException("External CSS file was not created.", cssPath);

        // -----------------------------------------------------------------
        // 7. Verify that the HTML references the external CSS file.
        // -----------------------------------------------------------------
        string htmlContent = File.ReadAllText(htmlPath);
        bool hasLinkTag = Regex.IsMatch(
            htmlContent,
            @"<link\s+rel=[""']stylesheet[""']\s+type=[""']text/css[""']\s+href=[""'].*styles\.css[""']",
            RegexOptions.IgnoreCase);

        if (!hasLinkTag)
            throw new InvalidOperationException("HTML does not contain a link to the external CSS file.");

        // -----------------------------------------------------------------
        // 8. Indicate success.
        // -----------------------------------------------------------------
        Console.WriteLine("HTML conversion completed successfully.");
        Console.WriteLine($"HTML file: {htmlPath}");
        Console.WriteLine($"CSS file: {cssPath}");
    }
}
