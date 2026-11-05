using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image based on the parameters from the DISPLAYBARCODE field.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Map the string barcode type to a SymbologyEncodeType using EncodeTypes static values.
        SymbologyEncodeType encodeType = EncodeTypes.Code128; // default

        if (!string.IsNullOrEmpty(parameters.BarcodeType))
        {
            string type = parameters.BarcodeType.ToUpperInvariant();
            if (type == "QR")
                encodeType = EncodeTypes.QR;
            else if (type == "EAN13")
                encodeType = EncodeTypes.EAN13;
            else if (type == "CODE128")
                encodeType = EncodeTypes.Code128;
            // Additional types can be added here as needed.
        }

        MemoryStream stream = new MemoryStream();
        using (BarcodeGenerator generator = new BarcodeGenerator(encodeType, parameters.BarcodeValue))
        {
            // Set barcode and background colors.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            // Save as PNG to the stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }

        stream.Position = 0;
        return stream;
    }

    // Legacy method delegates to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";
        field.ForegroundColor = "0x000000"; // Black
        field.BackgroundColor = "0xFFFFFF"; // White

        // Register the custom barcode generator so that barcodes are rendered as images.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to generate the barcode image.
        doc.UpdateFields();

        // Save the document as RTF; barcodes will appear as images.
        doc.Save("output.rtf");
    }
}
