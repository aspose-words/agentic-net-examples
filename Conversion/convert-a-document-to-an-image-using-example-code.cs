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
            // Path to the source Word document.
            string inputPath = @"C:\Docs\SampleDocument.docx";

            // Path where the rendered image will be saved.
            string outputPath = @"C:\Docs\SampleDocument.png";

            // Load the document from file using the Document(string) constructor.
            Document doc = new Document(inputPath);

            // Create ImageSaveOptions to specify image rendering settings.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render the first page only (default is all pages, but for a single image we limit it).
                PageSet = new PageSet(0),

                // Set the resolution (dots per inch). 300 DPI gives good quality.
                Resolution = 300,

                // Optional: set background color to transparent.
                PaperColor = System.Drawing.Color.Transparent
            };

            // Save the document as an image using the Save(string, SaveOptions) overload.
            doc.Save(outputPath, options);

            Console.WriteLine("Document has been converted to image successfully.");
        }
    }
}
