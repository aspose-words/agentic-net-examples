using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string pngPath = "sample.png";
        const string pdfPath = "output.pdf";

        // Create a simple PNG image using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the image with a solid color.
                graphics.Clear(Color.CornflowerBlue);
            }

            // Save the bitmap as a PNG file.
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // Verify that the PNG file was created.
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("Failed to create the PNG image.");

        // Create a new Word document and insert the PNG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);

        // Save the document as a PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF conversion was not successful.");

        // Clean up the temporary PNG file (optional).
        File.Delete(pngPath);
    }
}
