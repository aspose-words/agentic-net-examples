using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "Input.docx";

        // Path where the rendered image will be saved.
        string outputPath = "Output.png";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Create image save options for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
        // Example: set a higher resolution for better quality.
        options.Resolution = 300; // DPI

        // Save the first page of the document as an image.
        // When saving to an image format Aspose.Words renders only the first page.
        doc.Save(outputPath, options);
    }
}
