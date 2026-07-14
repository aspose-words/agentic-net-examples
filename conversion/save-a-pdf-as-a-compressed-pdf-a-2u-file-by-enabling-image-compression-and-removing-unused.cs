using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string inputPath = "sample.docx";
        const string outputPath = "compressed_pdfa2u.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample Word document with some text and an image.
        // -----------------------------------------------------------------
        Document createdDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createdDoc);
        builder.Writeln("This is a sample document for PDF/A‑2u conversion with image compression.");

        // Create a simple red square image using Aspose.Drawing (no System.Drawing usage)
        using (MemoryStream imageStream = new MemoryStream())
        {
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.Red);
                }

                // Save the bitmap as PNG into the memory stream
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image into the document from the byte array
                builder.InsertImage(imageStream.ToArray());
            }
        }

        // Save the created document as DOCX (bootstrap file)
        createdDoc.Save(inputPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the previously saved DOCX file.
        // -----------------------------------------------------------------
        Document docToConvert = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Configure PDF save options:
        //    - PDF/A‑2u compliance
        //    - JPEG image compression
        //    - JPEG quality (adjustable)
        //    - Optimize output to remove unused objects
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 70,               // Moderate compression
            OptimizeOutput = true           // Remove unused objects
        };

        // Save the document as a compressed PDF/A‑2u file
        docToConvert.Save(outputPath, pdfOptions);

        // -----------------------------------------------------------------
        // 4. Validation – ensure the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The PDF file '{outputPath}' was not created.");

        // Optional: report success (no interactive prompts required)
        Console.WriteLine($"PDF/A‑2u file saved successfully: {outputPath}");
    }
}
