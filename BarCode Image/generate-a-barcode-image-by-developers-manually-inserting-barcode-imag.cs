using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up barcode parameters for a QR code.
        BarcodeParameters barcodeParams = new BarcodeParameters
        {
            BarcodeType = "QR",
            BarcodeValue = "Sample123",
            BackgroundColor = "0xF8BD69",
            ForegroundColor = "0xB5413B",
            ErrorCorrectionLevel = "3",
            ScalingFactor = "250",
            SymbolHeight = "1000",
            SymbolRotation = "0"
        };

        // Generate the barcode image using the built‑in generator.
        using (Stream imgStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParams))
        {
            // Insert the image into the document with a target size (width x height) in points.
            // 1 point = 1/72 inch. Adjust the values to achieve the desired resolution.
            builder.InsertImage(imgStream, 150, 150);
        }

        // Save the document to a DOCX file.
        doc.Save("BarcodeDocument.docx");
    }
}
