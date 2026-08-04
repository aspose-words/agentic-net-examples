using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public class Program
{
    // Custom barcode generator that uses Aspose.BarCode to create QR code images.
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        // Generates a barcode image based on the supplied Word barcode parameters.
        public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
        {
            // Create a BarcodeGenerator for QR codes.
            var generator = new BarcodeGenerator(EncodeTypes.QR, parameters.BarcodeValue);

            // Save the generated barcode to a memory stream in PNG format.
            var stream = new MemoryStream();
            generator.Save(stream, BarCodeImageFormat.Png);
            stream.Position = 0;
            return stream;
        }

        // For legacy BARCODE fields – delegate to the primary method.
        public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
        {
            return GetBarcodeImage(parameters);
        }
    }

    public static void Main()
    {
        // Create a new empty Word document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Register the custom barcode generator (required for rendered outputs such as PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Insert a DISPLAYBARCODE field using the typed API.
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure the field to display a QR code.
        field.BarcodeType = "QR";
        field.BarcodeValue = "HelloWorld";

        // Update fields so that the barcode image is generated.
        doc.UpdateFields();

        // Save the document as PDF (the custom generator ensures the QR code is rendered).
        doc.Save("DisplayBarcodeQRCode.pdf");
    }
}
