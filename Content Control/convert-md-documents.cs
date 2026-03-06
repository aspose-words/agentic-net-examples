using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main(string[] args)
    {
        // Input Markdown file path.
        string inputPath = "input.md";

        // Desired output file path (PDF in this example).
        string outputPath = "output.pdf";

        // Open the Markdown file as a stream.
        using (FileStream inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            // Configure load options for Markdown.
            var loadOptions = new MarkdownLoadOptions
            {
                // Preserve empty lines from the source Markdown.
                PreserveEmptyLines = true
            };

            // Load the Markdown document into an Aspose.Words Document.
            Document doc = new Document(inputStream, loadOptions);

            // Save the document to the target format (PDF here).
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
