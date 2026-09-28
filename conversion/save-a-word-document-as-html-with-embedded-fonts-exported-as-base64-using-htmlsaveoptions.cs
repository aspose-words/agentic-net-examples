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
        builder.Font.Name = "Arial";
        builder.Writeln("Hello, this is a sample document with embedded fonts.");

        // Configure HTML save options. The default behavior embeds fonts as Base64 when possible.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

        string outputPath = "output.html";
        doc.Save(outputPath, saveOptions);

        // Validate that the HTML file was created and contains data.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("HTML output with embedded fonts was not created.");
        }

        Console.WriteLine($"HTML file saved successfully to '{outputPath}'.");
    }
}
