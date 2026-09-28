using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample JPEG image (1200x900) to be inserted into the document.
        const string sampleImagePath = "sample.jpg";
        CreateSampleJpeg(sampleImagePath, 1200, 900);

        // Create a Word document and insert the sample JPEG image twice.
        const string inputDocPath = "input.docx";
        CreateDocumentWithImages(inputDocPath, sampleImagePath);

        // Load the document for processing.
        Document doc = new Document(inputDocPath);

        // Collect all JPEG shapes.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        List<Shape> jpegShapes = new List<Shape>();
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
                jpegShapes.Add(shape);
        }

        if (jpegShapes.Count == 0)
            throw new InvalidOperationException("No JPEG images found in the document.");

        int index = 0;
        foreach (Shape shape in jpegShapes)
        {
            // Extract original image bytes.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load original bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    // Determine if resizing is needed.
                    if (originalBitmap.Width > 800)
                    {
                        int newWidth = 800;
                        int newHeight = (int)(originalBitmap.Height * (800.0 / originalBitmap.Width));

                        // Create resized bitmap.
                        using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                        {
                            using (Graphics g = Graphics.FromImage(resizedBitmap))
                            {
                                g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                            }

                            // Save resized bitmap to a memory stream as JPEG.
                            using (MemoryStream resizedStream = new MemoryStream())
                            {
                                resizedBitmap.Save(resizedStream, ImageFormat.Jpeg);
                                resizedStream.Position = 0;

                                // Replace image data in the shape.
                                shape.ImageData.SetImage(resizedStream);
                            }
                        }

                        // Optionally save the resized image to disk for verification.
                        string resizedImagePath = $"resized-{index}.jpg";
                        using (FileStream fileOut = new FileStream(resizedImagePath, FileMode.Create, FileAccess.Write))
                        {
                            shape.ImageData.Save(fileOut);
                        }
                    }
                }
            }
            index++;
        }

        // Save the modified document.
        const string outputDocPath = "output.docx";
        doc.Save(outputDocPath);
    }

    private static void CreateSampleJpeg(string path, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle for visual content.
                using (Pen pen = new Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 50, 50, width - 100, height - 100);
                }
            }
            bitmap.Save(path, ImageFormat.Jpeg);
        }
    }

    private static void CreateDocumentWithImages(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document with JPEG images:");
        builder.InsertImage(imagePath);
        builder.Writeln();
        builder.InsertImage(imagePath);
        doc.Save(docPath);
    }
}
