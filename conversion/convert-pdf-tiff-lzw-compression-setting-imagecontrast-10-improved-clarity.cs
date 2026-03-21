using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToTiffConverter
{
    static void Main()
    {
        const string pdfPath = "Input.pdf";
        const string tiffPath = "Output.tiff";

        // Ensure a source PDF exists. If not, create a simple one.
        if (!File.Exists(pdfPath))
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Sample PDF content for conversion.");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Load the source PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw
            // ImageContrast property is not supported for TIFF; omitted to avoid runtime error.
        };

        // Save the document as a TIFF image using the specified options.
        pdfDoc.Save(tiffPath, tiffOptions);
        Console.WriteLine($"Converted '{pdfPath}' to '{tiffPath}'.");
    }
}
