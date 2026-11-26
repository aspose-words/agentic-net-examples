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
        builder.Writeln("Hello, this is a sample document for HTML export.");

        // Define output HTML file path.
        string outputPath = "output.html";

        // Configure HTML save options.
        // The default behavior does not generate external JavaScript files,
        // so no additional property needs to be set.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Save the document as HTML with the specified options.
        doc.Save(outputPath, saveOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output HTML file was not created.");
        }

        // Ensure the file is not empty.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("Output HTML file is empty.");
        }
    }
}
