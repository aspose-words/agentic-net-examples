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
        const int width = 100;
        const int height = 100;
        const string sampleImagePath = "sample.png";

        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Aspose.Drawing.Color.White);
        using (SolidBrush brush = new SolidBrush(Aspose.Drawing.Color.Red))
        {
            graphics.FillRectangle(brush, 0, 0, width, height);
        }
        bitmap.Save(sampleImagePath);
        graphics.Dispose();
        bitmap.Dispose();

        // Create a Word document and insert the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Extract all PNG images from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the extracted image.
            string extractedPath = $"extracted-{imageIndex}.png";
            shape.ImageData.Save(extractedPath);

            // Load the extracted image for processing.
            using (Bitmap extractedBmp = new Bitmap(extractedPath))
            {
                // Prepare a bitmap for the processed image.
                using (Bitmap processedBmp = new Bitmap(extractedBmp.Width, extractedBmp.Height))
                {
                    using (Graphics g = Graphics.FromImage(processedBmp))
                    {
                        // Apply a hue rotation of 180 degrees (implemented as color inversion).
                        float[][] matrixElements = new float[][]
                        {
                            new float[] { -1, 0, 0, 0, 1 },
                            new float[] { 0, -1, 0, 0, 1 },
                            new float[] { 0, 0, -1, 0, 1 },
                            new float[] { 0, 0, 0, 1, 0 },
                            new float[] { 0, 0, 0, 0, 1 }
                        };
                        ColorMatrix colorMatrix = new ColorMatrix(matrixElements);
                        ImageAttributes attributes = new ImageAttributes();
                        attributes.SetColorMatrix(colorMatrix);

                        g.DrawImage(
                            extractedBmp,
                            new Rectangle(0, 0, processedBmp.Width, processedBmp.Height),
                            0,
                            0,
                            extractedBmp.Width,
                            extractedBmp.Height,
                            GraphicsUnit.Pixel,
                            attributes);
                    }

                    // Save the processed image.
                    string processedPath = $"processed-{imageIndex}.png";
                    processedBmp.Save(processedPath);
                    if (!File.Exists(processedPath))
                        throw new Exception($"Processed image not saved: {processedPath}");
                }
            }

            imageIndex++;
        }

        // Validate that at least one image was processed.
        if (imageIndex == 0)
            throw new Exception("No images were extracted from the document.");

        // Clean up temporary files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
