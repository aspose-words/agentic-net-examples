using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;   // Required for ShapeType

public class Program
{
    public static void Main()
    {
        // Define file and folder names.
        string baseDir = Directory.GetCurrentDirectory();
        string pdfPath = Path.Combine(baseDir, "sample.pdf");
        string htmlPath = Path.Combine(baseDir, "output.html");
        string imagesFolder = Path.Combine(baseDir, "Images");

        // Ensure the images folder exists for external image files.
        Directory.CreateDirectory(imagesFolder);

        // -----------------------------------------------------------------
        // Step 1: Create a simple PDF document using Aspose.Words.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Use a non‑standard font to guarantee that a font resource will be exported.
        builder.Font.Name = "Courier New";
        builder.Writeln("This is a sample PDF document with a non‑standard font.");

        // Add an image to demonstrate external image extraction.
        builder.InsertImage(ImageFromPlaceholder());

        // Save the document as PDF.
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the sample PDF file.");

        // -----------------------------------------------------------------
        // Step 2: Load the PDF and convert it to HTML.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportFontsAsBase64 = true,      // Embed fonts directly in the CSS as Base64.
            ExportFontResources = true,      // Ensure font resources are processed.
            ExportImagesAsBase64 = false,    // Keep images as external files.
            ImagesFolder = imagesFolder,     // Folder where external images will be saved.
            PrettyFormat = true
        };

        pdfDoc.Save(htmlPath, htmlOptions);

        // -----------------------------------------------------------------
        // Step 3: Validate the conversion results.
        // -----------------------------------------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        // Ensure at least one image file was written to the images folder.
        string[] imageFiles = Directory.GetFiles(imagesFolder);
        if (imageFiles.Length == 0)
            throw new InvalidOperationException("No image files were exported during HTML conversion.");

        // Confirm that fonts were embedded as Base64.
        string htmlContent = File.ReadAllText(htmlPath);
        if (!htmlContent.Contains("@font-face") || !htmlContent.Contains("base64"))
            throw new InvalidOperationException("Fonts were not embedded as Base64 in the HTML.");
    }

    // Helper method to generate a simple placeholder image stream.
    private static Stream ImageFromPlaceholder()
    {
        // Create a temporary document containing a rectangle shape.
        Document tempDoc = new Document();
        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
        tempBuilder.InsertShape(ShapeType.Rectangle, 100, 100);

        // Save the temporary document as a PNG image into a memory stream.
        using (MemoryStream ms = new MemoryStream())
        {
            tempDoc.Save(ms, SaveFormat.Png);
            ms.Position = 0;
            // Return a fresh stream containing the PNG data.
            return new MemoryStream(ms.ToArray());
        }
    }
}
