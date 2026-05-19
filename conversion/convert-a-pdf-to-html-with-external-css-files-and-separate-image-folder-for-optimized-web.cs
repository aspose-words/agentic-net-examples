using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document and save it as PDF (input for conversion).
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF file was not created.");

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Prepare output paths.
        string htmlPath = "output.html";
        string cssPath = "output.css";
        string imagesFolder = "images";

        // Ensure the images folder is clean and exists.
        if (Directory.Exists(imagesFolder))
            Directory.Delete(imagesFolder, true);
        Directory.CreateDirectory(imagesFolder);

        // Configure HTML save options for external CSS and separate images folder.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            CssStyleSheetType = CssStyleSheetType.External,
            CssStyleSheetFileName = cssPath,
            ImagesFolder = imagesFolder,
            ImagesFolderAlias = imagesFolder // Use a relative URI for the images.
        };

        // Save the PDF as HTML using the configured options.
        pdfDoc.Save(htmlPath, htmlOptions);

        // Validate that the conversion produced the expected files.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The HTML file was not created.");
        if (!File.Exists(cssPath))
            throw new InvalidOperationException("The external CSS file was not created.");
        if (!Directory.Exists(imagesFolder) || Directory.GetFiles(imagesFolder).Length == 0)
            throw new InvalidOperationException("No images were exported to the images folder.");

        // Indicate successful completion.
        Console.WriteLine("PDF successfully converted to HTML with external CSS and images folder.");
    }
}
