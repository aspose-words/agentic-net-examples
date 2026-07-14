using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ImageToPdfConverter
{
    public static void Main()
    {
        // Define file names for the sample images and the output PDF.
        const string pngPath = "sample_image.png";
        const string jpgPath = "sample_image.jpg";
        const string outputPdf = "combined_images.pdf";

        // Create a PNG image if it does not already exist.
        if (!File.Exists(pngPath))
        {
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.Red);
                }
                bitmap.Save(pngPath, ImageFormat.Png);
            }
        }

        // Create a JPEG image if it does not already exist.
        if (!File.Exists(jpgPath))
        {
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.Blue);
                }
                bitmap.Save(jpgPath, ImageFormat.Jpeg);
            }
        }

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the PNG image.
        builder.InsertImage(pngPath);
        builder.Writeln(); // Add a line break between images.

        // Insert the JPEG image.
        builder.InsertImage(jpgPath);
        builder.Writeln();

        // Save the document as a PDF file.
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF file was not created.");

        // Optionally, clean up the sample images.
        // File.Delete(pngPath);
        // File.Delete(jpgPath);
    }
}
