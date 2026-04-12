using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

namespace BatchPngToPdf
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Create sample PNG images using Aspose.Drawing.
            string[] imagePaths = new string[3];
            for (int i = 0; i < imagePaths.Length; i++)
            {
                string imageFile = Path.Combine(artifactsDir, $"SampleImage{i + 1}.png");
                CreateSamplePng(imageFile, i + 1);
                imagePaths[i] = imageFile;
            }

            // Build a Word document where each image occupies its own page.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 0; i < imagePaths.Length; i++)
            {
                // Insert the PNG image.
                builder.InsertImage(imagePaths[i]);

                // Add a page break after each image except the last one.
                if (i < imagePaths.Length - 1)
                {
                    builder.InsertBreak(BreakType.PageBreak);
                }
            }

            // Save the document as a single PDF file.
            string pdfPath = Path.Combine(artifactsDir, "CombinedImages.pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("The PDF file was not created.", pdfPath);

            Console.WriteLine($"PDF successfully created at: {pdfPath}");
        }

        // Generates a simple PNG image with a colored background and a label.
        private static void CreateSamplePng(string filePath, int index)
        {
            const int width = 400;
            const int height = 300;

            using (Bitmap bitmap = new Bitmap(width, height))
            {
                // Obtain a Graphics object for drawing.
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    // Fill background with a distinct color.
                    Aspose.Drawing.Color backgroundColor = index switch
                    {
                        1 => Aspose.Drawing.Color.LightBlue,
                        2 => Aspose.Drawing.Color.LightGreen,
                        _ => Aspose.Drawing.Color.LightCoral
                    };
                    graphics.Clear(backgroundColor);

                    // Draw a label in the center.
                    Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24);
                    Aspose.Drawing.Color textColor = Aspose.Drawing.Color.Black;
                    string text = $"Image {index}";

                    // Measure text size.
                    SizeF textSize = graphics.MeasureString(text, font);
                    float x = (width - textSize.Width) / 2;
                    float y = (height - textSize.Height) / 2;

                    using (SolidBrush brush = new SolidBrush(textColor))
                    {
                        graphics.DrawString(text, font, brush, x, y);
                    }
                }

                // Save the bitmap as PNG.
                bitmap.Save(filePath, ImageFormat.Png);
            }
        }
    }
}
