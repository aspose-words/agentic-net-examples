using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class TiffToPdfConverter
{
    static void Main()
    {
        // Paths to the source TIFF images.
        string[] tiffFiles = new string[]
        {
            @"C:\Images\Page1.tiff",
            @"C:\Images\Page2.tiff",
            @"C:\Images\Page3.tiff"
        };

        // Create a new empty Word document – this will become the PDF container.
        Document pdfDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDoc);

        // Insert each TIFF image as a full‑page picture.
        for (int i = 0; i < tiffFiles.Length; i++)
        {
            // Load the image file and insert it at the current cursor position.
            builder.InsertImage(tiffFiles[i]);

            // After each image except the last, insert a page break so the next image starts on a new page.
            if (i < tiffFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the assembled document as a PDF file.
        pdfDoc.Save(@"C:\Output\Combined.pdf", SaveFormat.Pdf);
    }
}
