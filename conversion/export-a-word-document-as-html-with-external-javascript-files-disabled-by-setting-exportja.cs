using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class ExportHtmlExample
{
    public static void Main()
    {
        // Create a new Word document and add sample content.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This is a sample Word document.");
        builder.Writeln("It will be exported to HTML without external JavaScript files.");

        // Configure HTML save options. By default, external JavaScript files are not generated.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Define the output HTML file path.
        string outputPath = "output.html";

        // Save the document as HTML using the configured options.
        source.Save(outputPath, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output HTML file was not created.");
        }

        // Optionally, read the file to ensure it contains content.
        string htmlContent = File.ReadAllText(outputPath);
        if (string.IsNullOrWhiteSpace(htmlContent))
        {
            throw new InvalidOperationException("The generated HTML file is empty.");
        }

        // Indicate successful completion.
        Console.WriteLine("HTML export completed successfully.");
    }
}
