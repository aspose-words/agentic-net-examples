using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchTiffToPng
{
    public static void Main()
    {
        // Prepare deterministic working directory
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -------------------------------------------------
        // Step 1: Create sample TIFF images with DPI metadata
        // -------------------------------------------------
        string[] tiffFiles =
        {
            Path.Combine(workDir, "sample1.tiff"),
            Path.Combine(workDir, "sample2.tiff")
        };

        for (int i = 0; i < tiffFiles.Length; i++)
        {
            // Create bitmap using Aspose.Drawing
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200);
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);

            // Draw deterministic content
            g.Clear(Aspose.Drawing.Color.White);
            Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20);
            g.DrawString($"Img {i + 1}", font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black), new Aspose.Drawing.PointF(20, 80));

            // Set DPI (e.g., 150x150) and save as TIFF
            bitmap.SetResolution(150f, 150f);
            bitmap.Save(tiffFiles[i], Aspose.Drawing.Imaging.ImageFormat.Tiff);

            // Clean up drawing resources
            g.Dispose();
            font.Dispose();
            bitmap.Dispose();
        }

        // -------------------------------------------------
        // Step 2: Insert TIFF images into a Word document
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        foreach (string tiffPath in tiffFiles)
        {
            // Ensure the file exists before insertion
            if (!File.Exists(tiffPath))
                throw new FileNotFoundException("TIFF file not found.", tiffPath);

            builder.InsertImage(tiffPath);
            builder.Writeln(); // separate images
        }

        string docPath = Path.Combine(workDir, "DocumentWithTiffs.docx");
        doc.Save(docPath);

        // -------------------------------------------------
        // Step 3: Load the document and convert each image to lossless PNG preserving DPI
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int pngCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue; // Skip shapes without images

            // Save the image data to a memory stream
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0; // Reset stream position before reading

                // Load the image into Aspose.Drawing.Bitmap
                using (Aspose.Drawing.Bitmap srcBitmap = new Aspose.Drawing.Bitmap(imgStream))
                {
                    // Preserve original DPI
                    float hDpi = srcBitmap.HorizontalResolution;
                    float vDpi = srcBitmap.VerticalResolution;

                    // Create a copy for PNG saving
                    using (Aspose.Drawing.Bitmap pngBitmap = new Aspose.Drawing.Bitmap(srcBitmap))
                    {
                        pngBitmap.SetResolution(hDpi, vDpi);

                        string pngPath = Path.Combine(workDir, $"converted_{pngCount + 1}.png");
                        pngBitmap.Save(pngPath, Aspose.Drawing.Imaging.ImageFormat.Png);

                        // Validate that the PNG file was created
                        if (!File.Exists(pngPath))
                            throw new InvalidOperationException($"Failed to create PNG file: {pngPath}");

                        pngCount++;
                    }
                }
            }
        }

        // Validation: ensure at least one PNG was created
        if (pngCount == 0)
            throw new InvalidOperationException("No images were found or converted.");

        // Optional cleanup (commented out for inspection)
        // foreach (string file in tiffFiles) File.Delete(file);
        // File.Delete(docPath);
    }
}
