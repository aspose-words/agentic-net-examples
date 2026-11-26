using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Minimal 1x1 transparent PNG.
    private static readonly byte[] BlankPng = new byte[]
    {
        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
        0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
        0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
        0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
        0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
        0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
        0x42,0x60,0x82
    };

    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Disable rendering for values prefixed with "DISABLE:".
        if (parameters.BarcodeValue != null && parameters.BarcodeValue.StartsWith("DISABLE:"))
        {
            var blankStream = new MemoryStream();
            blankStream.Write(BlankPng, 0, BlankPng.Length);
            blankStream.Position = 0;
            return blankStream;
        }

        // Normal barcode generation.
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a normal barcode field.
        FieldDisplayBarcode normalField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        normalField.BarcodeType = "CODE128";
        normalField.BarcodeValue = "1234567890";
        normalField.ForegroundColor = "0x000000";
        normalField.BackgroundColor = "0xFFFFFF";

        // Insert a barcode field that should be disabled in PDF.
        FieldDisplayBarcode disabledField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        disabledField.BarcodeType = "CODE128";
        disabledField.BarcodeValue = "DISABLE:987654321";
        disabledField.ForegroundColor = "0xFF0000";
        disabledField.BackgroundColor = "0xFFFFFF";

        // Update fields for DOCX.
        doc.UpdateFields();

        // Save as DOCX (both barcodes visible as fields).
        doc.Save("output.docx");

        // Register custom barcode generator for rendered outputs.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Re-update fields to apply custom generator.
        doc.UpdateFields();

        // Save as PDF (disabled barcode will appear blank).
        doc.Save("output.pdf");
    }
}
