using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputFolder = "input";
        string outputFolder = "output";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample TIFF images with different DPI values
        var tiffInfos = new List<(string FileName, float DpiX, float DpiY)>
        {
            ("image1.tif", 72f, 72f),
            ("image2.tif", 150f, 150f),
            ("image3.tif", 300f, 300f)
        };

        foreach (var info in tiffInfos)
        {
            string path = Path.Combine(inputFolder, info.FileName);
            using (Bitmap bmp = new Bitmap(200, 200))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                }
                bmp.SetResolution(info.DpiX, info.DpiY);
                bmp.Save(path, ImageFormat.Tiff);
            }
        }

        // Create a Word document and insert the TIFF images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (var info in tiffInfos)
        {
            string imgPath = Path.Combine(inputFolder, info.FileName);
            builder.InsertParagraph();
            builder.InsertImage(imgPath);
        }
        string docPath = "sample.docx";
        doc.Save(docPath);

        // Reload the document to simulate extraction scenario
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage) continue;

            // Save the original image to a memory stream
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Load the image with Aspose.Drawing
                using (Bitmap bitmap = new Bitmap(imgStream))
                {
                    // Preserve DPI
                    float dpiX = bitmap.HorizontalResolution;
                    float dpiY = bitmap.VerticalResolution;

                    // Prepare output PNG path
                    string pngPath = Path.Combine(outputFolder, $"converted-{imageIndex}.png");

                    // Ensure DPI is set before saving
                    bitmap.SetResolution(dpiX, dpiY);
                    bitmap.Save(pngPath, ImageFormat.Png);
                    imageIndex++;
                }
            }
        }

        // Validation
        string[] pngFiles = Directory.GetFiles(outputFolder, "*.png");
        if (pngFiles.Length == 0)
            throw new Exception("No PNG files were created during conversion.");

        // Output result summary (non-interactive)
        Console.WriteLine($"Converted {pngFiles.Length} TIFF images to PNG. Files are located in '{outputFolder}'.");
    }
}
