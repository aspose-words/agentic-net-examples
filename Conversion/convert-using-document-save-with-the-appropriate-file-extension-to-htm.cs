using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOCM file path
        string inputPath = "input.docm";

        // Output HTML file path (extension .html indicates HTML format)
        string outputPath = "output.html";

        // Load the DOCM document
        Document doc = new Document(inputPath);

        // Create HtmlSaveOptions specifying the HTML save format
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Save the document as HTML
        doc.Save(outputPath, saveOptions);
    }
}
