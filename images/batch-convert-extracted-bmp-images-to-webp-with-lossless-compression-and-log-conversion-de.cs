using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare deterministic file names
        string[] bmpFiles = { "image1.bmp", "image2.bmp" };
        string docFile = "sample.docx";

        // Create sample BMP images
        for (int i = 0; i < bmpFiles.Length; i++)
        {
            int width = 200;
            int height = 100;

            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
            Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap);
            try
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle with text to differentiate images
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 2))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }

                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 16))
                using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                {
                    g.DrawString($"Image {i + 1}", font, brush, new Aspose.Drawing.PointF(20, 40));
                }

                bitmap.Save(bmpFiles[i]);
            }
            finally
            {
                g.Dispose();
                bitmap.Dispose();
            }
        }

        // Create a Word document and insert the BMP images
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < bmpFiles.Length; i++)
        {
            // Insert image and obtain the created shape
            Shape shape = builder.InsertImage(bmpFiles[i]);
            // Store original file name in the shape's title for later reference
            shape.Title = Path.GetFullPath(bmpFiles[i]);
            // Add a line break after each image
            builder.Writeln();
        }
        doc.Save(docFile);

        // Load the document for extraction and conversion
        Document loadedDoc = new Document(docFile);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int conversionCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Retrieve original file information
            string originalPath = shape.Title ?? "unknown";
            long originalSize = originalPath != "unknown" && File.Exists(originalPath)
                ? new FileInfo(originalPath).Length
                : 0;

            // Prepare output WebP file name
            string outputFile = $"converted-{conversionCount + 1}.webp";

            // Save the image as WebP (lossless by default if supported)
            shape.ImageData.Save(outputFile);

            // Validate that the output file was created
            if (!File.Exists(outputFile))
                throw new InvalidOperationException($"Failed to create output file: {outputFile}");

            long newSize = new FileInfo(outputFile).Length;

            // Log conversion details
            Console.WriteLine($"Converted '{Path.GetFileName(originalPath)}' ({originalSize} bytes) to '{Path.GetFileName(outputFile)}' ({newSize} bytes).");

            conversionCount++;
        }

        // Ensure at least one conversion occurred
        if (conversionCount == 0)
            throw new InvalidOperationException("No images were converted.");

        // Clean up temporary files (optional)
        foreach (string bmp in bmpFiles)
        {
            if (File.Exists(bmp))
                File.Delete(bmp);
        }
        if (File.Exists(docFile))
            File.Delete(docFile);
    }
}
