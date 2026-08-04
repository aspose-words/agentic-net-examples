using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class PdfToHtmlConverter
{
    public static void Main()
    {
        // Prepare folders
        string outputFolder = Directory.GetCurrentDirectory();
        string imagesFolder = Path.Combine(outputFolder, "Images");
        Directory.CreateDirectory(imagesFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample PDF file (source) with some text and an image.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF document generated for conversion.");

        // Create a simple bitmap using Aspose.Drawing (no System.Drawing usage)
        string tempImagePath = Path.Combine(outputFolder, "sample.png");
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }
            bitmap.Save(tempImagePath);
        }

        // Insert the generated image into the document
        builder.InsertImage(tempImagePath);

        // Save the document as PDF
        string pdfPath = Path.Combine(outputFolder, "input.pdf");
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // ---------------------------------------------------------------
        // 2. Load the PDF and convert it to HTML with external CSS & images
        // ---------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export CSS to an external file
            CssStyleSheetType = CssStyleSheetType.External,
            CssStyleSheetFileName = Path.Combine(outputFolder, "output.css"),

            // Export images to a separate folder
            ImagesFolder = imagesFolder,
            ExportImagesAsBase64 = false,

            // Optional: keep the HTML tidy
            PrettyFormat = true
        };

        string htmlPath = Path.Combine(outputFolder, "output.html");
        pdfDoc.Save(htmlPath, htmlOptions);

        // ------------------------------
        // 3. Validate the conversion output
        // ------------------------------
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("HTML output file was not created.");

        if (!File.Exists(htmlOptions.CssStyleSheetFileName))
            throw new InvalidOperationException("External CSS file was not created.");

        if (!Directory.Exists(imagesFolder))
            throw new InvalidOperationException("Images folder was not created.");

        string[] imageFiles = Directory.GetFiles(imagesFolder);
        if (imageFiles.Length == 0)
            throw new InvalidOperationException("No images were exported to the images folder.");

        // Clean up temporary image used for source creation
        if (File.Exists(tempImagePath))
            File.Delete(tempImagePath);
    }
}
