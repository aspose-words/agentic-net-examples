using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

namespace ImageExtractionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a deterministic sample PNG image.
            const string sampleImagePath = "sample.png";
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
            Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
            graphics.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.SolidBrush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Red))
            {
                graphics.FillRectangle(brush, 10, 10, 80, 80);
            }
            bitmap.Save(sampleImagePath);
            graphics.Dispose();
            bitmap.Dispose();

            // Create a DOCX document and insert the sample image.
            const string docPath = "sample.docx";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(sampleImagePath);
            doc.Save(docPath);

            // Load the document for image extraction.
            Document loadedDoc = new Document(docPath);

            // Find the first shape that contains an image.
            Shape imageShape = null;
            NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    imageShape = shape;
                    break;
                }
            }

            if (imageShape == null)
                throw new InvalidOperationException("No image found in the document.");

            // Extract the image, convert to BMP, and place into a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                // Save the original image data to a stream.
                imageShape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image with Aspose.Drawing and re‑save as BMP.
                using (Aspose.Drawing.Image drawingImage = Aspose.Drawing.Image.FromStream(originalStream))
                using (MemoryStream bmpStream = new MemoryStream())
                {
                    drawingImage.Save(bmpStream, Aspose.Drawing.Imaging.ImageFormat.Bmp);
                    bmpStream.Position = 0; // Reset for downstream consumption.

                    if (bmpStream.Length == 0)
                        throw new InvalidOperationException("Extracted BMP image stream is empty.");

                    // Pass the BMP stream to the dummy API.
                    ProcessImageStream(bmpStream);
                }
            }

            // Clean up generated files (optional).
            File.Delete(sampleImagePath);
            File.Delete(docPath);
        }

        // Dummy API method that consumes the image stream.
        private static void ProcessImageStream(Stream imageStream)
        {
            // For demonstration, read the stream length.
            Console.WriteLine($"Image stream received with length: {imageStream.Length} bytes.");
            // Real API logic would go here.
        }
    }
}
