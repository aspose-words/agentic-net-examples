using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportWordToHtmlWithExternalCss
{
    public static void Main()
    {
        // Define file and folder names.
        string workDir = Directory.GetCurrentDirectory();
        string docPath = Path.Combine(workDir, "sample.docx");
        string htmlPath = Path.Combine(workDir, "sample.html");

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Hello Aspose.Words!");
        sourceDoc.Save(docPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the created document.
        // -----------------------------------------------------------------
        Document doc = new Document(docPath);

        // -----------------------------------------------------------------
        // 3. Configure HtmlFixedSaveOptions to disable embedded CSS.
        //    This causes the CSS to be saved as an external .css file.
        // -----------------------------------------------------------------
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedCss = false // Disable embedding => external CSS file.
        };

        // -----------------------------------------------------------------
        // 4. Save the document as HTML.
        // -----------------------------------------------------------------
        doc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the HTML file and the external CSS file were created.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // The external CSS is placed in a folder with the same name as the HTML file (without extension).
        string cssFolder = Path.Combine(workDir, Path.GetFileNameWithoutExtension(htmlPath));
        if (!Directory.Exists(cssFolder))
            throw new InvalidOperationException("CSS folder was not created.");

        // Find a .css file inside the folder.
        string[] cssFiles = Directory.GetFiles(cssFolder, "*.css");
        if (cssFiles.Length == 0)
            throw new InvalidOperationException("External CSS file was not created.");

        // Optional: verify that the HTML references the external CSS file.
        string htmlContent = File.ReadAllText(htmlPath);
        string cssFileName = Path.GetFileName(cssFiles[0]);
        string expectedLinkPattern = $"<link rel=\"stylesheet\" type=\"text/css\" href=\"{Path.GetFileNameWithoutExtension(htmlPath)}/{cssFileName}\"";
        if (!Regex.IsMatch(htmlContent, Regex.Escape(expectedLinkPattern), RegexOptions.IgnoreCase))
            throw new InvalidOperationException("HTML does not reference the external CSS file.");

        // All checks passed.
        Console.WriteLine("HTML and external CSS files were successfully created.");
    }
}
