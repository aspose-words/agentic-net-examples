using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputPath = "sample.docx";
        const string outputPath = "sample.html";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        // Add a hyperlink that could contain JavaScript (for demonstration).
        builder.InsertHyperlink("Click me", "javascript:alert('Hello')", false);
        doc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Save as HTML with external JavaScript disabled.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // This option removes JavaScript from links to improve security.
            RemoveJavaScriptFromLinks = true
        };
        loadedDoc.Save(outputPath, htmlOptions);

        // -----------------------------------------------------------------
        // 4. Validate that the HTML file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The expected HTML file '{outputPath}' was not created.");

        // Optional: Output a simple confirmation (no interactive input required).
        Console.WriteLine($"Document successfully saved to '{outputPath}' with JavaScript removed.");
    }
}
