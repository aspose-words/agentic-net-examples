using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page document and save it as PDF.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        string pdfPath = Path.Combine(outputFolder, "sample.pdf");
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the source PDF file.");

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Prepare image save options for PNG (lossless) output.
        // -----------------------------------------------------------------
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // PNG is lossless; we can still control resolution if desired.
            Resolution = 300,                     // 300 DPI for archival quality.
            ImageColorMode = ImageColorMode.None // Preserve original colors.
        };

        // -----------------------------------------------------------------
        // 4. Render each page of the PDF to a separate PNG file.
        // -----------------------------------------------------------------
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Render only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            string pngPath = Path.Combine(outputFolder, $"page_{pageIndex + 1}.png");
            pdfDoc.Save(pngPath, pngOptions);

            // Validate that the image was saved correctly.
            if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
                throw new InvalidOperationException($"Failed to save page {pageIndex + 1} as PNG.");
        }

        // -----------------------------------------------------------------
        // 5. Indicate successful completion.
        // -----------------------------------------------------------------
        Console.WriteLine($"PDF was split into {pdfDoc.PageCount} PNG images in folder: {outputFolder}");
    }
}
