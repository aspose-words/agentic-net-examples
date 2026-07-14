using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample PNG image using Aspose.Drawing.
        string sampleImagePath = Path.Combine(outputDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 20, 20, 160, 160);
                }
            }
            bitmap.Save(sampleImagePath);
        }

        // 2. Create a Word document, insert the sample image, and save it as PDF.
        string pdfPath = Path.Combine(outputDir, "document.pdf");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        // Save as PDF (default options are sufficient for this example).
        doc.Save(pdfPath, SaveFormat.Pdf);

        // 3. Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // 4. Extract all images from the PDF.
        NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // 4a. Save the extracted image to a temporary file (preserve original format).
            string tempImagePath = Path.Combine(outputDir, $"extracted_{imageIndex}.png");
            shape.ImageData.Save(tempImagePath);

            // 4b. Convert the extracted image to JPEG with 85% quality.
            string jpegPath = Path.Combine(outputDir, $"image_{imageIndex}.jpg");
            Document tempDoc = new Document();
            DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
            tempBuilder.InsertImage(tempImagePath);

            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 85
            };
            // Save only the first page (which contains the inserted image).
            jpegOptions.PageSet = new PageSet(0);
            tempDoc.Save(jpegPath, jpegOptions);

            // Validate that the JPEG file was created.
            if (!File.Exists(jpegPath))
                throw new InvalidOperationException($"Failed to create JPEG file: {jpegPath}");

            // Clean up the temporary extracted image.
            File.Delete(tempImagePath);

            imageIndex++;
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found in the PDF document.");

        // Optional: indicate completion (no interactive output required).
        // The program will exit automatically.
    }
}
