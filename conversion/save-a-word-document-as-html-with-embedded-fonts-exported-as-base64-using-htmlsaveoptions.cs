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
        builder.Writeln("This is a sample document with embedded fonts.");

        // Configure HTML save options.
        // Note: In some older versions of Aspose.Words the properties
        // ExportEmbeddedFonts and ExportEmbeddedFontsAsBase64 are not available.
        // If they are present, they can be set to true to embed fonts as Base64.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Uncomment the following lines if the properties exist in the referenced version:
        // htmlOptions.ExportEmbeddedFonts = true;
        // htmlOptions.ExportEmbeddedFontsAsBase64 = true;

        // Save the document as HTML with the specified options.
        string outputPath = "output.html";
        doc.Save(outputPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        // Ensure the file has content.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The HTML output file is empty.");
    }
}
