using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOT (Word template) document from file.
        Document doc = new Document("Template.dot");

        // Configure image save options – render pages as PNG images.
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Optional: set resolution, color mode, etc., if needed.
            // HorizontalResolution = 300,
            // VerticalResolution = 300,
            // ImageColorMode = ImageColorMode.None
        };

        // Save the first page of the document as an image.
        // To save all pages, iterate over doc.PageCount and adjust PageSet if required.
        doc.Save("Output.png", imageOptions);
    }
}
