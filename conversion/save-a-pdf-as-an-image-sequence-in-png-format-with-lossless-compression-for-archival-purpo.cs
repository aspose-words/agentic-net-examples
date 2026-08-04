using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two pages.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample PDF content - Page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Sample PDF content - Page 2.");

        // Save the document as PDF (input file for conversion).
        const string pdfPath = "input.pdf";
        source.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert each page of the PDF to a separate PNG image.
        for (int i = 0; i < pdfDoc.PageCount; i++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render the specific page.
                PageSet = new PageSet(i),

                // Use a high resolution for archival quality.
                Resolution = 300,

                // Ensure color mode is unchanged (PNG is lossless).
                ImageColorMode = ImageColorMode.None
            };

            string imagePath = $"output_page_{i + 1}.png";
            pdfDoc.Save(imagePath, options);

            // Validate that the image was created.
            if (!File.Exists(imagePath))
                throw new InvalidOperationException($"Failed to create image: {imagePath}");
        }

        // Clean up the temporary PDF if desired.
        // File.Delete(pdfPath);
    }
}
