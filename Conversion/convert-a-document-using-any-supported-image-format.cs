using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsImageConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source document (any format supported by Aspose.Words)
            string sourcePath = @"C:\Input\SampleDocument.docx";

            // Path to the output image file (choose any supported image format, e.g., PNG)
            string outputPath = @"C:\Output\SampleDocument.png";

            // Load the document from the file system.
            Document doc = new Document(sourcePath);

            // Configure image save options – specify the desired image format.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Png);

            // Optional: adjust image quality, resolution, etc. (commented out defaults)
            // imageOptions.Resolution = 300;          // DPI
            // imageOptions.JpegQuality = 90;         // For JPEG formats
            // imageOptions.ImageColorMode = ImageColorMode.None;

            // Save the first page of the document as an image.
            // When saving to an image format, only the first page is rendered.
            doc.Save(outputPath, imageOptions);

            Console.WriteLine("Document has been converted to image successfully.");
        }
    }
}
