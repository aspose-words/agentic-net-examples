using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing; // Required by the package list, though not used directly.

public class Program
{
    public static void Main()
    {
        // Create a sample Word document with multiple pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Page 1 - Sample content for PDF conversion.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 - Additional content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 - Final page.");

        // Save the document as PDF (the source for image extraction).
        const string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Prepare image save options for PNG (lossless compression).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Ensure each page is saved individually.
            // The PageSet will be overridden inside the loop for each page.
            ImageColorMode = ImageColorMode.None // Preserve original colors.
        };

        // Export each page of the PDF as a separate PNG file.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Set the page to render.
            pngOptions.PageSet = new PageSet(pageIndex);

            // Define the output file name.
            string outputFile = $"page_{pageIndex + 1}.png";

            // Save the current page as PNG.
            pdfDoc.Save(outputFile, pngOptions);

            // Validate that the file was created.
            if (!File.Exists(outputFile))
                throw new InvalidOperationException($"Expected image file '{outputFile}' was not created.");
        }

        // Clean up the temporary PDF file.
        if (File.Exists(pdfPath))
            File.Delete(pdfPath);
    }
}
