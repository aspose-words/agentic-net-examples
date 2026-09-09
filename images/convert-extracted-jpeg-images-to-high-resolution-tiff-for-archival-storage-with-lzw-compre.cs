using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample JPEG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bitmap))
        {
            g.Clear(Aspose.Drawing.Color.LightBlue);
            g.DrawEllipse(new Pen(Aspose.Drawing.Color.DarkRed, 5), 20, 20, 160, 160);
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the JPEG image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract JPEG images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Save the original JPEG to a memory stream.
            using (MemoryStream jpegStream = new MemoryStream())
            {
                shape.ImageData.Save(jpegStream);
                jpegStream.Position = 0; // Reset before reuse.

                // -----------------------------------------------------------------
                // 4. Create a temporary document containing the extracted image.
                // -----------------------------------------------------------------
                Document tempDoc = new Document();
                DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
                tempBuilder.InsertImage(jpegStream.ToArray());

                // -----------------------------------------------------------------
                // 5. Save the temporary document as a high‑resolution TIFF with LZW compression.
                // -----------------------------------------------------------------
                ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
                {
                    TiffCompression = TiffCompression.Lzw,
                    Resolution = 300 // High resolution (300 DPI).
                };

                string tiffPath = Path.Combine(artifactsDir, $"ExtractedImage_{imageIndex}.tiff");
                tempDoc.Save(tiffPath, tiffOptions);

                // Validate that the TIFF file was created.
                if (!File.Exists(tiffPath))
                    throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");

                imageIndex++;
            }
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
            throw new InvalidOperationException("No JPEG images were found to convert.");

        // The program finishes automatically.
    }
}
