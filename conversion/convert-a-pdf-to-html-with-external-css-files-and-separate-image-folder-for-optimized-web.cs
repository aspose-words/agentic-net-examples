using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output locations.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string pdfPath = Path.Combine(outputDir, "sample.pdf");
        string htmlPath = Path.Combine(outputDir, "sample.html");
        string cssPath = Path.Combine(outputDir, "styles.css");
        string imagesFolder = Path.Combine(outputDir, "Images");
        Directory.CreateDirectory(imagesFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample PDF document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document generated for conversion.");
        // Add an image to ensure image extraction works.
        // (Using a placeholder image generated as a simple PNG byte array.)
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
        using (MemoryStream imgStream = new MemoryStream(pngBytes))
        {
            builder.InsertImage(imgStream);
        }

        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF and convert it to HTML with external CSS and images folder.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External,
            CssStyleSheetFileName = cssPath,
            ImagesFolder = imagesFolder,
            // Ensure that image URIs in HTML point to the Images folder.
            ImagesFolderAlias = "Images"
        };

        pdfDoc.Save(htmlPath, saveOptions);

        // -----------------------------------------------------------------
        // 3. Validate the conversion results.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        if (!File.Exists(cssPath))
            throw new InvalidOperationException("External CSS file was not created.");

        if (!Directory.Exists(imagesFolder))
            throw new InvalidOperationException("Images folder was not created.");

        string[] imageFiles = Directory.GetFiles(imagesFolder);
        if (imageFiles.Length == 0)
            throw new InvalidOperationException("No images were extracted to the images folder.");

        // -----------------------------------------------------------------
        // 4. Indicate success (optional console output).
        // -----------------------------------------------------------------
        Console.WriteLine("PDF successfully converted to HTML.");
        Console.WriteLine($"HTML file: {htmlPath}");
        Console.WriteLine($"CSS file: {cssPath}");
        Console.WriteLine($"Images folder: {imagesFolder}");
    }
}
