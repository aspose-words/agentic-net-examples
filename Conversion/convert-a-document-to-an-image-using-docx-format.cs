using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the rendered image will be saved.
        string outputPath = "output.png";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Set up image save options to render the document as a PNG image.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);
        // Example: render only the first page.
        // saveOptions.PageSet = new PageSet(0);

        // Save the document pages as images.
        doc.Save(outputPath, saveOptions);
    }
}
