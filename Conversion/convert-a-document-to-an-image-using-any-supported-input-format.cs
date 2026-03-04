using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertToImage
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words)
        string inputPath = "input.docx";

        // Path where the rendered image will be saved
        string outputPath = "output.png";

        // Load the document using the Document constructor (load rule)
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for PNG format (create rule)
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
        // Example: set resolution to 300 DPI
        options.Resolution = 300;

        // Save the first page of the document as an image (save rule)
        doc.Save(outputPath, options);
    }
}
