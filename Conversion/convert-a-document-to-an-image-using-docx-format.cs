using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class DocxToImageConverter
{
    static void Main()
    {
        // Path to the folder that contains the input DOCX file.
        string docsFolder = @"C:\Docs\";

        // Input DOCX file name.
        string inputFile = "input.docx";

        // Output image file name (PNG format in this example).
        string outputFile = "output.png";

        // Load the DOCX document.
        Document doc = new Document(docsFolder + inputFile);

        // Create ImageSaveOptions to specify image format and rendering options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Render only the first page (zero‑based index). Remove this line to render all pages.
        options.PageSet = new PageSet(0);

        // Save the document as an image.
        doc.Save(docsFolder + outputFile, options);
    }
}
