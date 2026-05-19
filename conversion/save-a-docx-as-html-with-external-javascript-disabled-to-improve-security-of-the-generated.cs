using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SaveDocxAsHtml
{
    public static void Main()
    {
        // Define file names.
        const string docxPath = "sample.docx";
        const string htmlPath = "sample.html";

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        // Save the DOCX so that we can load it later (simulates an existing file).
        doc.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Configure HTML save options.
        //    - RemoveJavaScriptFromLinks disables JavaScript in links,
        //      improving the security of the generated HTML.
        // -----------------------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            RemoveJavaScriptFromLinks = true
        };

        // -----------------------------------------------------------------
        // 4. Save the document as HTML using the configured options.
        // -----------------------------------------------------------------
        loadedDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // 5. Verify that the HTML file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML file was not created.");

        // Optional: output a confirmation (no interactive input required).
        Console.WriteLine("DOCX successfully converted to HTML with JavaScript removed.");
    }
}
