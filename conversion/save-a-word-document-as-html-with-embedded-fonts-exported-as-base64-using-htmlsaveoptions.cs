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
        builder.Writeln("This is a sample document with embedded fonts.");

        // Configure HTML save options.
        // Note: In some older versions of Aspose.Words the properties
        // ExportEmbeddedFonts and ExportEmbeddedFontsAsBase64 are not available.
        // If they are present, they can be set here to embed fonts as Base64.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Uncomment the following lines if the Aspose.Words version supports them:
        // htmlOptions.ExportEmbeddedFonts = true;
        // htmlOptions.ExportEmbeddedFontsAsBase64 = true;

        // Define output HTML file path.
        string outputPath = "output.html";

        // Save the document as HTML with the specified options.
        doc.Save(outputPath, htmlOptions);

        // Validate that the HTML file was created and contains data.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The HTML output file was not created.");

        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The HTML output file is empty.");

        // Indicate success.
        Console.WriteLine("HTML file saved successfully.");
    }
}
