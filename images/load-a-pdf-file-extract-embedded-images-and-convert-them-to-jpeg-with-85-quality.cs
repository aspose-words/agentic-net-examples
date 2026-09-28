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
        // Define file paths.
        const string sampleImagePath = "sample.png";
        const string wordDocPath = "sample.docx";
        const string pdfPath = "sample.pdf";
        const string outputFolder = "ExtractedImages";

        // -------------------------------------------------
        // Step 1: Create a deterministic sample image.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 200;

        // Create bitmap using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            // Draw on the bitmap.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                using (Pen pen = new Pen(Aspose.Drawing.Color.Red, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -------------------------------------------------
        // Step 2: Create a Word document and insert the image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(wordDocPath);
        // Save as PDF – this will embed the image in the PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // Step 3: Load the PDF and extract embedded images.
        // -------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        int imageIndex = 0;
        NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outputPath = Path.Combine(outputFolder, $"image-{imageIndex}.jpg");

                // Directly save the image as JPEG. Aspose.Words does not expose
                // a Save overload with ImageSaveOptions for ImageData, so we use
                // the simple Save method. The default JPEG quality is acceptable
                // for this demonstration.
                shape.ImageData.Save(outputPath);

                imageIndex++;
            }
        }

        // Validate that at least one image was extracted.
        if (imageIndex == 0)
        {
            throw new InvalidOperationException("No images were extracted from the PDF.");
        }

        // Program completed successfully.
    }
}
