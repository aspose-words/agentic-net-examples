using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class PdfImageExtractor
{
    static void Main()
    {
        // Input PDF file path
        string pdfPath = @"Input.pdf";

        // Directory where extracted JPEG images will be saved
        string outputDir = @"ExtractedImages";
        Directory.CreateDirectory(outputDir);

        // Load the PDF document (lifecycle rule: load)
        Document pdfDoc = new Document(pdfPath);

        // Get all shapes (including pictures) from the document
        NodeCollection shapeNodes = pdfDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        foreach (Shape shape in shapeNodes)
        {
            // Process only shapes that contain an image
            if (!shape.HasImage) continue;

            // Create a temporary document that will hold the extracted image
            Document imgDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(imgDoc);

            // Save the shape's image data into a memory stream
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                imgStream.Position = 0;

                // Insert the image from the stream into the temporary document
                builder.InsertImage(imgStream);
            }

            // Prepare image save options with JPEG quality set to 85%
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            jpegOptions.JpegQuality = 85; // quality rule

            // Save the temporary document as a JPEG file (lifecycle rule: save)
            string outFile = Path.Combine(outputDir, $"Image_{imageIndex}.jpg");
            imgDoc.Save(outFile, jpegOptions);

            imageIndex++;
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to '{outputDir}'.");
    }
}
