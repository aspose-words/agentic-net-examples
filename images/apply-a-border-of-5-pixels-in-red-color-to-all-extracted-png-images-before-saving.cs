using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample PNG image.
        const string sampleImagePath = "sample.png";
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillEllipse(Brushes.Black, 25, 25, 50, 50);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // Create a document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "DocumentWithImage.docx";
        doc.Save(docPath);

        // Extract PNG images, apply a red 5‑pixel border, and save them.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int savedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            ImageData imgData = shape.ImageData;
            if (imgData.ImageType != ImageType.Png)
                continue;

            // Load the original image into a bitmap.
            using (MemoryStream ms = new MemoryStream())
            {
                imgData.Save(ms);
                ms.Position = 0;
                using (Bitmap original = new Bitmap(ms))
                {
                    int borderSize = 5;
                    int newWidth = original.Width + borderSize * 2;
                    int newHeight = original.Height + borderSize * 2;

                    // Create a new bitmap with a red background.
                    using (Bitmap bordered = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics g = Graphics.FromImage(bordered))
                        {
                            g.Clear(Color.Red);
                            g.DrawImage(
                                original,
                                new Rectangle(borderSize, borderSize, original.Width, original.Height));
                        }

                        string outPath = $"extracted-{++savedCount}.png";
                        bordered.Save(outPath, ImageFormat.Png);
                    }
                }
            }
        }

        if (savedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and saved.");

        // Clean up the sample files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
