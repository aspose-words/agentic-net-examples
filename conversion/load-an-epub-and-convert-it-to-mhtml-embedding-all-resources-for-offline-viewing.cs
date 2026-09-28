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
        // Create a temporary PNG image using Aspose.Drawing.
        string imagePath = "sample.png";
        CreateSampleImage(imagePath);

        // Build a sample document that contains the image.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample EPUB document with an embedded image.");
        builder.InsertImage(imagePath);
        builder.Writeln("End of document.");

        // Save the document as EPUB – this will be the input for conversion.
        string epubPath = "input.epub";
        sourceDoc.Save(epubPath, SaveFormat.Epub);

        // Load the EPUB document.
        Document epubDoc = new Document(epubPath);

        // Convert the EPUB to MHTML. Images are embedded automatically.
        string mhtmlPath = "output.mhtml";
        epubDoc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Validate that the MHTML file was created and is not empty.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
        {
            throw new InvalidOperationException("MHTML output was not created or is empty.");
        }

        // Optional cleanup – uncomment if you want to delete temporary files.
        // File.Delete(imagePath);
        // File.Delete(epubPath);
    }

    private static void CreateSampleImage(string path)
    {
        // Create a 100x100 pixel bitmap.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Obtain a Graphics object from the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with light blue.
                graphics.Clear(Color.LightBlue);

                // Draw a red ellipse.
                using (Pen pen = new Pen(Color.Red, 3))
                {
                    graphics.DrawEllipse(pen, 10, 10, 80, 80);
                }
            }

            // Save the bitmap to a PNG file.
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
