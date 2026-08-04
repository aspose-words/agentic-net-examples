using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractVideoThumbnails
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(workDir);
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample thumbnail image (PNG) using Aspose.Drawing.
        string thumbnailPath = Path.Combine(workDir, "thumb.png");
        const int imgWidth = 200;
        const int imgHeight = 150;
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                // Draw a simple rectangle to make the image recognizable.
                using (Pen pen = new Pen(Aspose.Drawing.Color.DarkBlue, 3))
                {
                    g.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }
            // Save the bitmap as PNG.
            bitmap.Save(thumbnailPath, ImageFormat.Png);
        }

        // 2. Create a DOCX document and insert the thumbnail image.
        //    In a real scenario this image would be the video thumbnail stored by Word.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with a video thumbnail (simulated by an image):");
        builder.InsertImage(thumbnailPath);
        string docPath = Path.Combine(workDir, "SampleWithVideo.docx");
        doc.Save(docPath);

        // 3. Load the document that supposedly contains video thumbnails.
        Document loadedDoc = new Document(docPath);

        // 4. Extract all images (thumbnails) from the document and save them as PNG files.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Ensure the image is saved as PNG regardless of its original type.
                string outFile = Path.Combine(outputDir, $"Thumbnail_{imageIndex}.png");
                // If the image is already PNG we can use Save(string). Otherwise, convert via stream.
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0;
                    using (Bitmap bmp = new Bitmap(ms))
                    {
                        bmp.Save(outFile, ImageFormat.Png);
                    }
                }
                imageIndex++;
            }
        }

        // 5. Validate that at least one thumbnail was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No thumbnail images were extracted from the document.");

        // The program finishes automatically; extracted PNG files are located in the Output folder.
    }
}
