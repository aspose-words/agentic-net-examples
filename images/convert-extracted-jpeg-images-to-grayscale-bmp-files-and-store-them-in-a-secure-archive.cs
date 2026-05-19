using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string imagesDir = Path.Combine(artifactsDir, "Images");
        Directory.CreateDirectory(imagesDir);

        // 1. Create a deterministic sample JPEG image.
        string jpegPath = Path.Combine(imagesDir, "sample.jpg");
        using (Bitmap bmp = new Bitmap(200, 200))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Aspose.Drawing.Color.White);
            // Draw a simple red ellipse.
            using (Brush brush = new SolidBrush(Aspose.Drawing.Color.Red))
            {
                g.FillEllipse(brush, 20, 20, 160, 160);
            }
            bmp.Save(jpegPath, ImageFormat.Jpeg);
        }

        // 2. Create a Word document and insert the JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        string docPath = Path.Combine(artifactsDir, "Document.docx");
        doc.Save(docPath);

        // 3. Extract JPEG images, convert each to grayscale BMP, and collect output paths.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        var bmpFiles = new System.Collections.Generic.List<string>();

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Convert to grayscale using the GrayScale property.
                shape.ImageData.GrayScale = true;

                // Save as BMP.
                string bmpPath = Path.Combine(imagesDir, $"image_{imageIndex}.bmp");
                shape.ImageData.Save(bmpPath);
                bmpFiles.Add(bmpPath);
                imageIndex++;
            }
        }

        // Validate that at least one BMP file was produced.
        if (bmpFiles.Count == 0)
            throw new InvalidOperationException("No JPEG images were found to convert.");

        // 4. Store the grayscale BMP files in a zip archive (simple secure container).
        string zipPath = Path.Combine(artifactsDir, "GrayscaleImages.zip");
        using (FileStream zipToOpen = new FileStream(zipPath, FileMode.Create))
        using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create))
        {
            foreach (string bmpFile in bmpFiles)
            {
                string entryName = Path.GetFileName(bmpFile);
                archive.CreateEntryFromFile(bmpFile, entryName);
            }
        }

        // Cleanup: optional deletion of intermediate files can be added here.
        // The example finishes without waiting for user input.
    }
}
