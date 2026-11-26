using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph.");
        builder.InsertParagraph();
        builder.Writeln("Another line with **bold** text.");

        // Configure HTML save options to add a CSS class name prefix.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            CssClassNamePrefix = "myPrefix_"
        };

        // Define the output HTML file path.
        string outputPath = "output.html";

        // Save the document as HTML using the configured options.
        doc.Save(outputPath, htmlOptions);

        // Validate that the HTML file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The HTML output file was not created.");
        }

        // Optionally, you could read the file to ensure it contains data.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The HTML output file is empty.");
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("HTML file saved successfully with CSS class prefix.");
    }
}
