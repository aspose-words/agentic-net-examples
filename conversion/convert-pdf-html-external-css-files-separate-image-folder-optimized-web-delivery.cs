using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToHtmlConverter
{
    static void Main()
    {
        // Create a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "PdfToHtmlDemo");
        Directory.CreateDirectory(workDir);

        // Paths for the source PDF, output HTML, CSS, and images folder.
        string pdfPath = Path.Combine(workDir, "sample.pdf");
        string htmlPath = Path.Combine(workDir, "sample.html");
        string cssFile = Path.Combine(workDir, "sample.css");
        string imagesFolder = Path.Combine(workDir, "images");

        // Ensure the images folder exists and is empty.
        if (Directory.Exists(imagesFolder))
            Directory.Delete(imagesFolder, true);
        Directory.CreateDirectory(imagesFolder);

        // Create a simple PDF document if it does not already exist.
        if (!File.Exists(pdfPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words PDF to HTML conversion!");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            CssStyleSheetType = CssStyleSheetType.External,
            CssStyleSheetFileName = cssFile,
            ImagesFolder = imagesFolder,
            ExportImagesAsBase64 = false,
            PrettyFormat = true
        };

        // Save the document as HTML using the configured options.
        pdfDoc.Save(htmlPath, saveOptions);

        Console.WriteLine($"HTML saved to: {htmlPath}");
        Console.WriteLine($"CSS saved to: {cssFile}");
        Console.WriteLine($"Images saved to: {imagesFolder}");
    }
}
