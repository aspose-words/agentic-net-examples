using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DotToImageConverter
{
    static void Main()
    {
        // Path to the source DOT (Word template) file.
        string inputPath = @"C:\Docs\Template.dot";

        // Path to the destination image file (PNG format in this example).
        string outputPath = @"C:\Docs\TemplateImage.png";

        // Load the DOT document using the Document(string) constructor.
        Document doc = new Document(inputPath);

        // Save the first page of the document as an image.
        // The Save method determines the format from the file extension,
        // but we explicitly specify the format to ensure correct handling.
        doc.Save(outputPath, SaveFormat.Png);
    }
}
