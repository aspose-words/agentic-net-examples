using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToPngConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\input.docx";

        // Path where the PNG image will be saved.
        string outputFile = @"C:\Docs\output.png";

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // Create image save options for PNG format.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // Optional: set resolution (dpi) if higher quality is required.
        // pngOptions.Resolution = 300;

        // Save the first page of the document as a PNG image.
        doc.Save(outputFile, pngOptions);
    }
}
