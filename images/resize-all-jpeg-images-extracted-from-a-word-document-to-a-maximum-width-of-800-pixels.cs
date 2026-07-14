using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample JPEG image (1200x900) using Aspose.Drawing
        string sampleJpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(1200, 900))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
            }
            bitmap.Save(sampleJpegPath, ImageFormat.Jpeg);
        }

        // 2. Create a Word document and insert the JPEG image multiple times
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleJpegPath);
        builder.InsertParagraph();
        builder.InsertImage(sampleJpegPath);
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        doc.Save(inputDocPath);

        // 3. Load the document (optional, we already have it)
        Document loadedDoc = new Document(inputDocPath);

        // 4. Extract JPEG images, resize if width > 800px, and save resized versions
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int jpegIndex = 0;
        int resizedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save original image to a memory stream
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load image with Aspose.Drawing
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;

                    // Determine if resizing is needed
                    if (originalWidth <= 800)
                    {
                        // No resizing needed, just save the original
                        string outPath = Path.Combine(artifactsDir, $"Image_{jpegIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
                        using (FileStream fs = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                        {
                            originalBitmap.Save(fs, ImageFormat.Jpeg);
                        }
                        resizedCount++;
                    }
                    else
                    {
                        // Calculate new dimensions while preserving aspect ratio
                        int newWidth = 800;
                        int newHeight = (int)Math.Round((double)originalHeight * newWidth / originalWidth);

                        using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                        {
                            using (Graphics g = Graphics.FromImage(resizedBitmap))
                            {
                                g.Clear(Color.Transparent);
                                g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                            }

                            string outPath = Path.Combine(artifactsDir, $"Resized_{jpegIndex}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
                            using (FileStream fs = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                            {
                                resizedBitmap.Save(fs, ImageFormat.Jpeg);
                            }
                            resizedCount++;
                        }
                    }
                }
            }

            jpegIndex++;
        }

        // Validation: ensure at least one resized image was produced
        if (resizedCount == 0)
            throw new InvalidOperationException("No JPEG images were processed.");

        // Optional: clean up sample files (comment out if you want to inspect them)
        // File.Delete(sampleJpegPath);
        // File.Delete(inputDocPath);
    }
}
