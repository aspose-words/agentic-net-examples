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
        // Prepare folders.
        string artifactsDir = "Artifacts";
        string archiveDir = Path.Combine(artifactsDir, "Archive");
        Directory.CreateDirectory(artifactsDir);
        Directory.CreateDirectory(archiveDir);

        // -----------------------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill with a solid color and draw a simple shape.
            graphics.Clear(Color.Blue);
            graphics.FillEllipse(new SolidBrush(Color.Yellow), 50, 50, 100, 100);
            // Save as JPEG – format is explicitly specified.
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Insert the JPEG image into a Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithJpeg.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Reload the document and extract JPEG images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        var jpegShapes = shapeNodes
            .OfType<Shape>()
            .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Jpeg)
            .ToList();

        if (!jpegShapes.Any())
            throw new InvalidOperationException("No JPEG images were found in the document.");

        int index = 0;
        foreach (Shape shape in jpegShapes)
        {
            // -----------------------------------------------------------------
            // 4. Convert the image to grayscale and save as BMP.
            // -----------------------------------------------------------------
            using (MemoryStream imgStream = new MemoryStream())
            {
                // Save the original image bytes to a stream.
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the image into a bitmap.
                using (Bitmap srcBmp = new Bitmap(imgStream))
                using (Bitmap grayBmp = new Bitmap(srcBmp.Width, srcBmp.Height))
                using (Graphics g = Graphics.FromImage(grayBmp))
                {
                    // Convert each pixel to grayscale.
                    for (int y = 0; y < srcBmp.Height; y++)
                    {
                        for (int x = 0; x < srcBmp.Width; x++)
                        {
                            Color pixel = srcBmp.GetPixel(x, y);
                            int gray = (int)(pixel.R * 0.3 + pixel.G * 0.59 + pixel.B * 0.11);
                            Color grayColor = Color.FromArgb(gray, gray, gray);
                            grayBmp.SetPixel(x, y, grayColor);
                        }
                    }

                    // Save the grayscale bitmap as BMP.
                    string bmpFileName = Path.Combine(archiveDir, $"image_{index}.bmp");
                    grayBmp.Save(bmpFileName, ImageFormat.Bmp);

                    if (!File.Exists(bmpFileName))
                        throw new InvalidOperationException($"Failed to create BMP file: {bmpFileName}");
                }
            }

            index++;
        }

        // -----------------------------------------------------------------
        // 5. Indicate successful completion.
        // -----------------------------------------------------------------
        Console.WriteLine($"Extracted and converted {index} image(s) to grayscale BMP files in '{archiveDir}'.");
    }
}
