using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

namespace SupportedImageFormatsForDoc
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add some content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Document demonstrating supported image formats when saved as DOC.");

            // Insert a sample JPEG image.
            // (Replace the path with an actual image file on your system.)
            string jpegPath = @"Images\sample.jpg";
            builder.InsertImage(jpegPath);

            // Save the document in the legacy DOC format using DocSaveOptions.
            // This ensures the document is saved with the .doc extension and Word 97‑2007 compatibility.
            DocSaveOptions docSaveOptions = new DocSaveOptions(SaveFormat.Doc);
            doc.Save("SupportedImages.doc", docSaveOptions);

            // -----------------------------------------------------------------
            // OPTIONAL: Render each page of the document to the image formats
            // that are supported when the document is later saved as DOC.
            // This demonstrates how to generate those images using ImageSaveOptions.
            // -----------------------------------------------------------------

            // Define the list of image formats supported for embedding in DOC.
            SaveFormat[] imageFormats = new SaveFormat[]
            {
                SaveFormat.Jpeg,   // JPEG
                SaveFormat.Png,    // PNG
                SaveFormat.Bmp,    // BMP
                SaveFormat.Gif,    // GIF
                SaveFormat.Emf,    // EMF
                SaveFormat.WebP    // WebP
            };

            // Render each page to each image format and save to disk.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                foreach (SaveFormat imgFormat in imageFormats)
                {
                    // Configure image save options for the current format.
                    ImageSaveOptions imgOptions = new ImageSaveOptions(imgFormat)
                    {
                        // Render only the current page.
                        PageSet = new PageSet(pageIndex)
                    };

                    // Build a filename like "Page0_Jpeg.png".
                    string extension = FileFormatUtil.SaveFormatToExtension(imgFormat);
                    string outFile = $"Page{pageIndex}_{imgFormat}{extension}";

                    // Save the rendered page as an image.
                    doc.Save(outFile, imgOptions);
                }
            }

            // -----------------------------------------------------------------
            // OPTIONAL: Create a new document and embed the generated images.
            // This shows that the images can be re‑inserted into a DOC file.
            // -----------------------------------------------------------------

            Document docWithImages = new Document();
            DocumentBuilder imgBuilder = new DocumentBuilder(docWithImages);
            imgBuilder.Writeln("Document containing images rendered in all supported formats:");

            foreach (SaveFormat imgFormat in imageFormats)
            {
                // Use the first page image for demonstration.
                string extension = FileFormatUtil.SaveFormatToExtension(imgFormat);
                string imagePath = $"Page0_{imgFormat}{extension}";

                if (File.Exists(imagePath))
                {
                    imgBuilder.InsertParagraph();
                    imgBuilder.InsertImage(imagePath);
                    imgBuilder.Writeln($"{imgFormat} image inserted.");
                }
            }

            // Save the final document as DOC.
            docWithImages.Save("AllSupportedImages.doc", new DocSaveOptions(SaveFormat.Doc));
        }
    }
}
