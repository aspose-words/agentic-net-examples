using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class TiffToPdfConverter
{
    public static void Main()
    {
        // Directory for temporary files
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create sample TIFF images
        int imageCount = 3;
        string[] tiffFiles = new string[imageCount];
        for (int i = 0; i < imageCount; i++)
        {
            string tiffPath = Path.Combine(workDir, $"sample{i + 1}.tiff");
            CreateSampleTiff(tiffPath, i + 1);
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF image: {tiffPath}");
            tiffFiles[i] = tiffPath;
        }

        // Create a new Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert each TIFF image on a separate page
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            // Insert the image; the method returns the created Shape
            Shape imgShape = builder.InsertImage(tiffFiles[i]);

            // Validate that the shape actually contains an image
            if (!imgShape.HasImage)
                throw new InvalidOperationException($"Inserted shape does not contain an image for file: {tiffFiles[i]}");

            // Add a page break after each image except the last one
            if (i < tiffFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Embed document metadata
        doc.BuiltInDocumentProperties.Title = "TIFF to PDF Example";
        doc.BuiltInDocumentProperties.Author = "Aspose Example";
        doc.BuiltInDocumentProperties.Subject = "Conversion of extracted TIFF images to PDF";
        doc.BuiltInDocumentProperties.Keywords = "TIFF,PDF,Aspose.Words";

        // Save the document as PDF
        string pdfPath = Path.Combine(workDir, "ConvertedImages.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);

        // Verify that the PDF was created
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"PDF file was not created: {pdfPath}");

        // Clean up temporary TIFF files (optional)
        foreach (string file in tiffFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }

        // Successful completion (no console output required)
    }

    // Creates a deterministic sample TIFF image with simple graphics
    private static void CreateSampleTiff(string filePath, int index)
    {
        int width = 200;
        int height = 200;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            // Obtain a Graphics object from the bitmap
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with a distinct color per index
                Aspose.Drawing.Color bgColor = (index % 2 == 0) ? Aspose.Drawing.Color.LightBlue : Aspose.Drawing.Color.LightGreen;
                graphics.Clear(bgColor);

                // Draw a simple rectangle border
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black, 3))
                {
                    graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }

                // Draw the index number in the center
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48, Aspose.Drawing.FontStyle.Bold))
                using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkRed))
                {
                    string text = index.ToString();
                    SizeF textSize = graphics.MeasureString(text, font);
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;
                    graphics.DrawString(text, font, brush, x, y);
                }
            }

            // Save as TIFF
            bitmap.Save(filePath);
        }
    }
}
