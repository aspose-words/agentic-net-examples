using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the Word template in DOT format.
        // The Document constructor automatically detects the format from the file extension.
        Document doc = new Document("InputTemplate.dot");

        // Create image save options for PNG format.
        // ImageSaveOptions derives from FixedPageSaveOptions and allows us to specify
        // rendering parameters such as resolution, page range, etc.
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Png);
        // Example: render at 300 DPI for higher quality.
        imageOptions.Resolution = 300;

        // Save the first page of the document as an image.
        // The Save method with (string, SaveOptions) follows the required lifecycle rule.
        doc.Save("OutputImage.png", imageOptions);
    }
}
