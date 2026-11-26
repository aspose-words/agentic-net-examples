using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample PDF document.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample PDF content for conversion.");
        string pdfPath = "sample.pdf";
        sourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure TIFF save options with LZW compression, increased contrast,
        // and limit the output to the first page.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw,
            ImageContrast = 1.0f, // Valid range is 0..1; using maximum contrast.
            PageSet = new PageSet(0) // Save only the first page.
        };

        // Save the PDF as a TIFF image.
        string tiffPath = "output.tiff";
        pdfDoc.Save(tiffPath, tiffOptions);

        // Validate that the TIFF file was created and is not empty.
        if (!File.Exists(tiffPath) || new FileInfo(tiffPath).Length == 0)
        {
            throw new InvalidOperationException("The TIFF file was not created or is empty.");
        }

        // Optional cleanup of the temporary PDF.
        // File.Delete(pdfPath);
    }
}
