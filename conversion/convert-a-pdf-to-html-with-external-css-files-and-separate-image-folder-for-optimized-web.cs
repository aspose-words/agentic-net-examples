using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Working directory
        string workDir = Directory.GetCurrentDirectory();

        // Create a simple PNG image using Aspose.Drawing
        string imagePath = Path.Combine(workDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.LightBlue);
            }
            bitmap.Save(imagePath);
        }

        // Build a Word document that contains some text and the image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document converted from PDF to HTML.");
        builder.InsertImage(imagePath);

        // Save the document as PDF (simulating the source PDF)
        string pdfPath = Path.Combine(workDir, "sample.pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Convert the document to HTML with an external CSS file and separate images folder
        string htmlPath = Path.Combine(workDir, "output.html");
        string cssPath = Path.Combine(workDir, "styles.css");
        string imagesFolder = Path.Combine(workDir, Path.GetFileNameWithoutExtension(htmlPath) + "_files");

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            CssStyleSheetType = CssStyleSheetType.External,
            CssStyleSheetFileName = Path.GetFileName(cssPath), // only file name is required
            ExportImagesAsBase64 = false,                     // ensure images are saved to folder
            ImagesFolder = imagesFolder,                      // explicit images folder
            ImagesFolderAlias = Path.GetFileName(imagesFolder) // folder name used in HTML
        };

        doc.Save(htmlPath, htmlOptions);

        // Validation
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        if (!File.Exists(cssPath))
            throw new InvalidOperationException("External CSS file was not created.");

        if (!Directory.Exists(imagesFolder))
            throw new InvalidOperationException("Images folder was not created.");

        string[] imageFiles = Directory.GetFiles(imagesFolder);
        if (imageFiles.Length == 0)
            throw new InvalidOperationException("No images were exported to the images folder.");

        // Optional cleanup (commented out)
        // File.Delete(imagePath);
        // File.Delete(pdfPath);
        // File.Delete(htmlPath);
        // File.Delete(cssPath);
        // Directory.Delete(imagesFolder, true);
    }
}
