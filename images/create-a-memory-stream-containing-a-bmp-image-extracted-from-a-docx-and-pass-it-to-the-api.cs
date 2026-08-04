using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample image using Aspose.Drawing
        const int width = 200;
        const int height = 200;
        const string sampleImagePath = "sample.png";

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white
                graphics.Clear(Aspose.Drawing.Color.White);
                // Draw a red ellipse
                using (Pen pen = new Pen(Aspose.Drawing.Color.Red, 5))
                {
                    graphics.DrawEllipse(pen, 10, 10, width - 20, height - 20);
                }
            }
            // Save the image to a local file
            bitmap.Save(sampleImagePath);
        }

        // Verify that the sample image was created
        if (!File.Exists(sampleImagePath))
            throw new FileNotFoundException("Failed to create the sample image.", sampleImagePath);

        // Step 2: Create a DOCX document and insert the image
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Verify that the document was saved
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to save the DOCX document.", docPath);

        // Step 3: Load the document and extract the image as BMP into a MemoryStream
        Document loadedDoc = new Document(docPath);
        Shape imageShape = null;

        foreach (Node node in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            Shape shape = (Shape)node;
            if (shape.HasImage)
            {
                imageShape = shape;
                break; // Assuming only one image for this example
            }
        }

        if (imageShape == null)
            throw new InvalidOperationException("No image shape found in the document.");

        // Prepare a memory stream for the BMP image
        using (MemoryStream bmpStream = new MemoryStream())
        {
            // Convert the original image bytes to BMP and write to the stream
            byte[] originalBytes = imageShape.ImageData.ImageBytes;
            using (MemoryStream srcStream = new MemoryStream(originalBytes))
            {
                using (Image img = Image.FromStream(srcStream))
                {
                    img.Save(bmpStream, ImageFormat.Bmp);
                }
            }

            // Reset stream position for subsequent reading
            bmpStream.Position = 0;

            // Optional: write the extracted BMP to a file for validation
            const string extractedBmpPath = "extracted.bmp";
            using (FileStream fileStream = new FileStream(extractedBmpPath, FileMode.Create, FileAccess.Write))
            {
                bmpStream.CopyTo(fileStream);
            }

            // Validate that the BMP file was created
            if (!File.Exists(extractedBmpPath))
                throw new FileNotFoundException("Failed to write the extracted BMP image.", extractedBmpPath);

            // Reset stream again before using it further
            bmpStream.Position = 0;

            // Step 4: Simulate passing the stream to an API by converting to Base64 and creating JSON payload
            byte[] bmpBytes = bmpStream.ToArray();
            string base64Image = Convert.ToBase64String(bmpBytes);
            var payload = new { ImageBase64 = base64Image };
            string json = JsonConvert.SerializeObject(payload, Formatting.Indented);

            // Output the JSON payload (could be sent to an API)
            Console.WriteLine(json);
        }

        // Clean up temporary files (optional)
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
        // File.Delete("extracted.bmp");
    }
}
