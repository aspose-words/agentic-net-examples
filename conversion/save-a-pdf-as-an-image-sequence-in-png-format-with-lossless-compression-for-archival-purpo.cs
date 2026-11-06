using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample Word document and save it as PDF (the source PDF).
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample PDF page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is a sample PDF page 2.");

        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Prepare PNG image save options (lossless compression is default for PNG).
        // -----------------------------------------------------------------
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // -----------------------------------------------------------------
        // 4. Save each page of the PDF as a separate PNG image.
        // -----------------------------------------------------------------
        string outputPattern = "output_page_{0}.png";

        int pageCount = pdfDoc.PageCount;
        for (int i = 0; i < pageCount; i++)
        {
            // Restrict the save operation to a single page.
            pngOptions.PageSet = new PageSet(i);

            string fileName = string.Format(outputPattern, i);
            pdfDoc.Save(fileName, pngOptions);
        }

        // -----------------------------------------------------------------
        // 5. Verify that PNG files were created for each page.
        // -----------------------------------------------------------------
        for (int i = 0; i < pageCount; i++)
        {
            string fileName = string.Format(outputPattern, i);
            if (!File.Exists(fileName))
                throw new InvalidOperationException($"Expected PNG file '{fileName}' was not created.");
        }
    }
}
