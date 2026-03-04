using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the Markdown file will be saved.
        string outputPath = "output.md";

        // Load the DOCX document using the default load options.
        Document doc = new Document(inputPath);

        // Create default Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

        // Explicitly set the format to Markdown (optional, default is already Markdown).
        saveOptions.SaveFormat = SaveFormat.Markdown;

        // Save the document as a Markdown file using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
