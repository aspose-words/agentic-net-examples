using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Define file paths
        string imagePath = Path.Combine(outputDir, "sample.png");
        string mhtmlPath = Path.Combine(outputDir, "sample.mhtml");
        string pdfPath = Path.Combine(outputDir, "sample.pdf");

        // Create a simple PNG image using Aspose.Drawing
        CreateSampleImage(imagePath);

        // Create a Word document, add text and the image, then save as MHTML
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with an embedded image.");
        builder.InsertImage(imagePath);
        doc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Load the MHTML file with proper LoadOptions
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFormat = LoadFormat.Mhtml; // specify that the source is MHTML
        Document mhtmlDoc = new Document(mhtmlPath, loadOptions);

        // Convert the loaded document to PDF
        mhtmlDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF file exists and is not empty
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
        }

        long pdfSize = new FileInfo(pdfPath).Length;
        if (pdfSize == 0)
        {
            throw new InvalidOperationException("PDF file is empty.");
        }

        // Confirmation output
        Console.WriteLine($"PDF successfully created at: {pdfPath} (size {pdfSize} bytes)");
    }

    private static void CreateSampleImage(string filePath)
    {
        // Create a 100x100 pixel bitmap
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Obtain a Graphics object from the bitmap
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with light blue
                graphics.Clear(Color.LightBlue);

                // Draw a red ellipse
                using (Pen pen = new Pen(Color.Red, 3))
                {
                    graphics.DrawEllipse(pen, 10, 10, 80, 80);
                }
            }

            // Save the bitmap as PNG
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
