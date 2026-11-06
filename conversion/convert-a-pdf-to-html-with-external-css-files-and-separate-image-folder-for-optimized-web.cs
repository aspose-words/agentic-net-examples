using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Hello PDF to HTML conversion using Aspose.Words.");
        string pdfPath = "sample.pdf";
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The sample PDF was not created.");

        // Load the PDF for conversion.
        Document pdfDocument = new Document(pdfPath);

        // Configure HTML save options for external CSS and separate images folder.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.CssStyleSheetType = CssStyleSheetType.External;
        htmlOptions.CssStyleSheetFileName = "styles.css";
        htmlOptions.ImagesFolder = "images";
        htmlOptions.ImagesFolderAlias = "images";
        htmlOptions.ExportImagesAsBase64 = false;

        // Perform the conversion.
        string htmlPath = "output.html";
        pdfDocument.Save(htmlPath, htmlOptions);

        // Validate that the HTML file, CSS file, and images folder were created.
        if (!File.Exists(htmlPath))
            throw new InvalidOperationException("The output HTML file was not created.");

        if (!File.Exists("styles.css"))
            throw new InvalidOperationException("The external CSS file was not created.");

        if (!Directory.Exists("images"))
            throw new InvalidOperationException("The images folder was not created.");

        // Example completed successfully.
    }
}
