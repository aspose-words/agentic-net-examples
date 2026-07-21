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
        builder.Writeln("Hello, Aspose.Words!");

        // Save the sample document as DOCX to simulate an existing file.
        string inputPath = "input.docx";
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the document from the file.
        Document loadedDoc = new Document(inputPath);

        // Configure HTML save options. External JavaScript files are disabled by default,
        // so no additional property needs to be set.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Save the document as HTML.
        string outputPath = "output.html";
        loadedDoc.Save(outputPath, htmlOptions);

        // Verify that the HTML file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("HTML output file was not created.");
        }

        // Optional: display a confirmation message.
        Console.WriteLine($"HTML file successfully created at: {Path.GetFullPath(outputPath)}");
    }
}
