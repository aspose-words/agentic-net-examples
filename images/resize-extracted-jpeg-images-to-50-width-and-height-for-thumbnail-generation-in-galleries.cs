using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample JPEG image.
        const string sampleImagePath = "sample.jpg";
        const int originalWidth = 200;
        const int originalHeight = 200;

        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Jpeg);
        }

        // Verify the sample image was created.
        if (!File.Exists(sampleImagePath))
            throw new Exception("Failed to create the sample JPEG image.");

        // Step 2: Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "document.docx";
        doc.Save(docPath);

        // Verify the document was saved.
        if (!File.Exists(docPath))
            throw new Exception("Failed to save the Word document.");

        // Step 3: Extract JPEG images from the document.
        List<string> extractedImagePaths = new List<string>();
        int imageIndex = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage)
            {
                string extractedPath = $"extracted-{imageIndex}.jpg";
                shape.ImageData.Save(extractedPath);
                extractedImagePaths.Add(extractedPath);
                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (extractedImagePaths.Count == 0)
            throw new Exception("No images were extracted from the document.");

        // Step 4: Resize each extracted JPEG to 50% for thumbnail generation.
        int thumbIndex = 0;
        foreach (string extractedPath in extractedImagePaths)
        {
            using (Bitmap originalBitmap = new Bitmap(extractedPath))
            {
                int thumbWidth = originalBitmap.Width / 2;
                int thumbHeight = originalBitmap.Height / 2;

                using (Bitmap thumbBitmap = new Bitmap(thumbWidth, thumbHeight))
                {
                    using (Graphics g = Graphics.FromImage(thumbBitmap))
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(originalBitmap, 0, 0, thumbWidth, thumbHeight);
                    }

                    string thumbPath = $"thumb-{thumbIndex}.jpg";
                    thumbBitmap.Save(thumbPath, ImageFormat.Jpeg);

                    // Validate thumbnail creation.
                    if (!File.Exists(thumbPath))
                        throw new Exception($"Thumbnail was not saved: {thumbPath}");

                    Console.WriteLine($"Thumbnail created: {thumbPath}");
                }
            }
            thumbIndex++;
        }

        // Cleanup: optional removal of intermediate files (comment out if inspection needed).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
        // foreach (var path in extractedImagePaths) File.Delete(path);
    }
}
