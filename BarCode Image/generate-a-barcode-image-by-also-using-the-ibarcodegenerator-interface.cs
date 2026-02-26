using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Returns a simple 1x1 PNG image for any barcode request.
    public Stream GetBarcodeImage(BarcodeParameters parameters)
    {
        return GetPlaceholderImage();
    }

    // Returns a simple 1x1 PNG image for old‑fashioned barcode fields.
    public Stream GetOldBarcodeImage(BarcodeParameters parameters)
    {
        return GetPlaceholderImage();
    }

    private Stream GetPlaceholderImage()
    {
        // Minimal PNG (1×1 transparent pixel) byte array.
        byte[] png = new byte[]
        {
            137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82,
            0,0,0,1,0,0,0,1,8,6,0,0,0,31,21,196,
            137,0,0,0,12,73,68,65,84,8,153,99,0,1,0,0,
            5,0,1,13,10,2,0,0,0,0,73,69,78,68,174,66,
            96,130
        };
        return new MemoryStream(png);
    }
}

class Program
{
    static void Main()
    {
        // Path to the source DOCX containing barcode fields.
        string inputPath = "Barcodes.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = "Barcodes.pdf";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Assign the custom barcode generator to the document.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields so that barcode images are generated.
        doc.UpdateFields();

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
