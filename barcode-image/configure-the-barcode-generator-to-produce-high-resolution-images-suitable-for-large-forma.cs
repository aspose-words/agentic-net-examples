using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a high‑resolution barcode image based on the provided parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Stream that will hold the PNG image.
        MemoryStream stream = new MemoryStream();

        // Map the field's BarcodeType string to a SymbologyEncodeType.
        SymbologyEncodeType symbology = EncodeTypes.Code128; // Default to Code128.
        if (!string.IsNullOrEmpty(parameters.BarcodeType))
        {
            string type = parameters.BarcodeType.ToUpperInvariant();
            if (type == "QR")
                symbology = EncodeTypes.QR;
            else if (type == "EAN13")
                symbology = EncodeTypes.EAN13;
            // Additional mappings can be added here if needed.
        }

        // Generate the barcode. The default image size is suitable for high‑resolution PDF printing.
        using (BarcodeGenerator generator = new BarcodeGenerator(symbology, parameters.BarcodeValue))
        {
            // Set barcode and background colors.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            // Save the barcode as PNG to the memory stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }

        // Reset the stream position for reading.
        stream.Position = 0;
        return stream;
    }

    // Legacy method – delegates to the current implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a typed DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890123";
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to render the barcode.
        doc.UpdateFields();

        // Save the document as a high‑resolution PDF.
        doc.Save("HighResolutionBarcode.pdf");
    }
}
