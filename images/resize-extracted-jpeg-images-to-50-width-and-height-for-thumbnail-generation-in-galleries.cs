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
        // Prepare a deterministic folder for all generated files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // 1. Create a sample JPEG image (200x200) using Aspose.Drawing.
        string inputImagePath = Path.Combine(dataDir, "input.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(inputImagePath, ImageFormat.Jpeg);
        }

        // 2. Insert the image into a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = Path.Combine(dataDir, "doc_with_image.docx");
        doc.Save(docPath);

        // 3. Load the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Save the shape's image to a memory stream.
                using (MemoryStream ms = new MemoryStream())
                {
                    shape.ImageData.Save(ms);
                    ms.Position = 0; // Reset before reading.

                    // Load the original image.
                    using (Bitmap original = new Bitmap(ms))
                    {
                        int newWidth = original.Width / 2;
                        int newHeight = original.Height / 2;

                        // Create a resized bitmap (50% of original size).
                        using (Bitmap resized = new Bitmap(newWidth, newHeight))
                        {
                            using (Graphics g = Graphics.FromImage(resized))
                            {
                                g.Clear(Color.White);
                                g.DrawImage(original, new Rectangle(0, 0, newWidth, newHeight));
                            }

                            // Save the thumbnail.
                            string thumbPath = Path.Combine(dataDir, $"thumb_{imageIndex}.jpg");
                            resized.Save(thumbPath, ImageFormat.Jpeg);
                        }
                    }
                }

                imageIndex++;
            }
        }

        // Validation: ensure at least one thumbnail was created.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were extracted and resized.");
    }
}
