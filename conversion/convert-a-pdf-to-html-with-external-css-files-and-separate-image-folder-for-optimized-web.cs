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
        // Prepare file and folder names
        string workDir = Directory.GetCurrentDirectory();
        string pdfPath = Path.Combine(workDir, "sample.pdf");
        string htmlPath = Path.Combine(workDir, "output.html");
        string cssPath = Path.Combine(workDir, "styles.css");
        string imagesFolder = Path.Combine(workDir, "images");
        string imageFilePath = Path.Combine(workDir, "sample.png");

        // Ensure clean state
        if (File.Exists(pdfPath)) File.Delete(pdfPath);
        if (File.Exists(htmlPath)) File.Delete(htmlPath);
        if (File.Exists(cssPath)) File.Delete(cssPath);
        if (Directory.Exists(imagesFolder)) Directory.Delete(imagesFolder, true);
        if (File.Exists(imageFilePath)) File.Delete(imageFilePath);

        // 1. Create a simple PNG image using Aspose.Drawing
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
                graphics.FillEllipse(new SolidBrush(Color.Yellow), 10, 10, 80, 80);
            }
            bitmap.Save(imageFilePath, ImageFormat.Png);
        }

        // 2. Create a Word document, insert the image, and save as PDF
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF generated from a Word document.");
        builder.InsertImage(imageFilePath);
        doc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // 3. Load the PDF back into Aspose.Words
        Document pdfDoc = new Document(pdfPath);

        // 4. Configure HTML save options for external CSS and separate images folder
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            CssStyleSheetFileName = cssPath,               // Full path to external CSS file
            CssStyleSheetType = CssStyleSheetType.External, // Ensure external CSS is generated
            ImagesFolder = imagesFolder,                   // Full path for images folder
            ImagesFolderAlias = "images",                  // Relative path used in HTML
            PrettyFormat = true
        };

        // 5. Save as HTML
        pdfDoc.Save(htmlPath, saveOptions);

        // 6. Validate output files and folders
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML file was not created.");

        if (!File.Exists(cssPath))
            throw new InvalidOperationException("External CSS file was not created.");

        if (!Directory.Exists(imagesFolder))
            throw new InvalidOperationException("Images folder was not created.");

        string[] imageFiles = Directory.GetFiles(imagesFolder);
        if (imageFiles.Length == 0)
            throw new InvalidOperationException("No images were exported to the images folder.");

        // Example completed successfully
        Console.WriteLine("PDF to HTML conversion completed successfully.");
    }
}
