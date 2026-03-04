using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToImageConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\Sample.docx";

        // Path where the rendered image will be saved.
        string outputFile = @"C:\Docs\Sample.png";

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // Save the first page of the document as a PNG image.
        // The Save method overload (string, SaveFormat) determines the format from the extension.
        doc.Save(outputFile, SaveFormat.Png);
    }
}
