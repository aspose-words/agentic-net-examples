using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");

        // Configure HTML save options.
        // Aspose.Words does not generate external JavaScript files by default,
        // so no additional property is required to disable them.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Save the document as HTML.
        string htmlPath = Path.Combine(outputDir, "Sample.html");
        doc.Save(htmlPath, saveOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML file was not created.");

        // Optional validation: ensure no <script> tags are present.
        string content = File.ReadAllText(htmlPath);
        if (content.Contains("<script", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("JavaScript was not disabled as expected.");
    }
}
