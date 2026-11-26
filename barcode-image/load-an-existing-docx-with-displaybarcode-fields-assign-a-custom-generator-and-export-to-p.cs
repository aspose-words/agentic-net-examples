using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image based on the parameters supplied by Aspose.Words.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Map the Word barcode type string to a SymbologyEncodeType used by Aspose.BarCode.
        SymbologyEncodeType encodeType = MapBarcodeType(parameters.BarcodeType);

        MemoryStream stream = new MemoryStream();
        using (BarcodeGenerator generator = new BarcodeGenerator(encodeType, parameters.BarcodeValue))
        {
            // Optional: set barcode colors.
            generator.Parameters.Barcode.BarColor = Aspose.Drawing.Color.Black;
            generator.Parameters.BackColor = Aspose.Drawing.Color.White;

            // Save the generated barcode as PNG into the stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }

        // Reset stream position for reading.
        stream.Position = 0;
        return stream;
    }

    // Legacy method – simply forwards to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }

    // Maps the string value from the DISPLAYBARCODE field to a SymbologyEncodeType.
    private static SymbologyEncodeType MapBarcodeType(string barcodeType)
    {
        if (string.IsNullOrEmpty(barcodeType))
            return EncodeTypes.Code128; // Default fallback.

        switch (barcodeType.Trim().ToUpperInvariant())
        {
            case "CODE128":
                return EncodeTypes.Code128;
            case "QR":
                return EncodeTypes.QR;
            case "CODE39":
                return EncodeTypes.Code39;
            case "EAN13":
                return EncodeTypes.EAN13;
            default:
                // Fallback to Code128 for unsupported types.
                return EncodeTypes.Code128;
        }
    }
}

public class Program
{
    public static void Main()
    {
        const string inputPath = "input.docx";
        const string outputPath = "output.pdf";

        Document doc;

        // If the input DOCX does not exist, create a simple document with a DISPLAYBARCODE field.
        if (!File.Exists(inputPath))
        {
            doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a typed DISPLAYBARCODE field.
            FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            field.BarcodeType = "CODE128";
            field.BarcodeValue = "1234567890";
            // Optional color settings (hex string format as required by the field).
            field.ForegroundColor = "0x000000";
            field.BackgroundColor = "0xFFFFFF";

            doc.UpdateFields();

            // Save the temporary DOCX so that the rest of the workflow works with a file.
            doc.Save(inputPath);
        }
        else
        {
            // Load the existing DOCX containing DISPLAYBARCODE fields.
            doc = new Document(inputPath);
        }

        // Register the custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Ensure all fields are up‑to‑date so that barcode images are generated.
        doc.UpdateFields();

        // Export the document to PDF.
        doc.Save(outputPath);
    }
}
