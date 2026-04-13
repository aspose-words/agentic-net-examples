using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Folder for generated files (current directory is sufficient)
        string[] bmpFiles = { "sample0.bmp", "sample1.bmp", "sample2.bmp" };

        // 1. Create deterministic BMP sample images using Aspose.Drawing
        for (int i = 0; i < bmpFiles.Length; i++)
        {
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            graphics.Clear(Aspose.Drawing.Color.White);
            // Draw a simple rectangle to make each image distinct
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2))
            {
                graphics.DrawRectangle(pen, 10, 10, 80, 80);
            }
            bitmap.Save(bmpFiles[i], ImageFormat.Bmp);
            graphics.Dispose();
            bitmap.Dispose();
        }

        // 2. Insert the BMP images into a Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string bmpPath in bmpFiles)
        {
            // InsertImage returns a Shape that automatically gets appended to the paragraph
            builder.InsertImage(bmpPath);
        }
        string docPath = "DocumentWithBmp.docx";
        doc.Save(docPath);

        // 3. Load the document and extract images, converting BMP to lossless PNG
        Document loadedDoc = new Document(docPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        string checksumFile = "checksums.txt";
        // Ensure checksum file starts empty
        File.WriteAllText(checksumFile, string.Empty);

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the image data to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading

                // Load the image with Aspose.Drawing to enable format conversion
                using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(imageStream))
                {
                    string pngPath = $"image{imageIndex}.png";
                    img.Save(pngPath, ImageFormat.Png); // Lossless PNG

                    // Compute SHA256 checksum of the PNG file
                    byte[] pngBytes = File.ReadAllBytes(pngPath);
                    using (SHA256 sha = SHA256.Create())
                    {
                        byte[] hash = sha.ComputeHash(pngBytes);
                        string hashString = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                        File.AppendAllText(checksumFile, $"{pngPath}: {hashString}{Environment.NewLine}");
                    }
                }
            }

            imageIndex++;
        }

        // 4. Validation – ensure at least one image was processed
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted and converted.");

        // Optional: output result paths (no console interaction required)
        // The program ends here.
    }
}
