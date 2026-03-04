using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ImagePageSetExample
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the rendered image will be saved.
        string outputPath = @"C:\Docs\RenderedPages.jpeg";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for JPEG format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);

        // Specify which pages to render.
        // The PageSet constructor accepts zero‑based page indices.
        // Here we render pages 1 and 2 (indices 0 and 1).
        saveOptions.PageSet = new PageSet(0, 1);

        // Save the selected pages as an image.
        doc.Save(outputPath, saveOptions);
    }
}
