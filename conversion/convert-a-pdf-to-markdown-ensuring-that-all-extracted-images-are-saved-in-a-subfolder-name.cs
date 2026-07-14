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
        // Define file and folder paths.
        string currentDir = Directory.GetCurrentDirectory();
        string pdfPath = Path.Combine(currentDir, "input.pdf");
        string markdownPath = Path.Combine(currentDir, "output.md");
        string assetsFolder = Path.Combine(currentDir, "assets");
        string imagePath = Path.Combine(currentDir, "sample.png");

        // Ensure the assets folder exists.
        Directory.CreateDirectory(assetsFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample image using Aspose.Drawing (no System.Drawing).
        // -----------------------------------------------------------------
        using (Bitmap bitmap = new Bitmap(200, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Resolve the ambiguous Font type by using the fully qualified name.
                Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 20);
                graphics.DrawString("Sample", font, Brushes.Black, new PointF(10, 40));
                font.Dispose();
            }
            bitmap.Save(imagePath, ImageFormat.Png);
        }

        // --------------------------------------------------------------
        // 2. Create a PDF document that contains some text and the image.
        // --------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF with an image.");
        builder.InsertImage(imagePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        // --------------------------------------------------------------
        // 3. Load the PDF and convert it to Markdown, extracting images.
        // --------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);
        MarkdownSaveOptions mdOptions = new MarkdownSaveOptions
        {
            ImagesFolder = assetsFolder // Images will be saved here.
        };
        pdfDoc.Save(markdownPath, mdOptions);

        // ------------------------------
        // 4. Validation of the results.
        // ------------------------------
        if (!File.Exists(markdownPath))
            throw new InvalidOperationException("The Markdown file was not created.");

        if (!Directory.Exists(assetsFolder) || Directory.GetFiles(assetsFolder).Length == 0)
            throw new InvalidOperationException("No images were extracted to the assets folder.");

        // Optional: output a short confirmation (no interactive input required).
        Console.WriteLine("Conversion completed successfully.");
        Console.WriteLine($"Markdown file: {markdownPath}");
        Console.WriteLine($"Extracted images folder: {assetsFolder}");
    }
}
