using System;
using Aspose.Words;
using Aspose.Words.Fields;

public sealed class Program
{
    // Inserts a DISPLAYBARCODE field with the specified settings.
    private static void InsertBarcode(DocumentBuilder builder, string barcodeType, string barcodeValue,
                                      string foregroundColor, string backgroundColor,
                                      string errorCorrectionLevel, string scalingFactor)
    {
        // Insert a typed DISPLAYBARCODE field and obtain the FieldDisplayBarcode object.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = barcodeType;               // e.g., "QR", "CODE128", "CODE39", "EAN13"
        field.BarcodeValue = barcodeValue;             // Text to encode.
        field.ForegroundColor = foregroundColor;      // Hex string, e.g., "0x000000".
        field.BackgroundColor = backgroundColor;      // Hex string, e.g., "0xFFFFFF".
        field.ErrorCorrectionLevel = errorCorrectionLevel; // For QR codes, optional.
        field.ScalingFactor = scalingFactor;           // Percentage as string, e.g., "100".
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a QR code.
        InsertBarcode(builder, "QR", "https://example.com",
                      "0x000000", "0xFFFFFF", "3", "100");
        builder.Writeln(); // Add a line break.

        // Insert a Code 128 barcode.
        InsertBarcode(builder, "CODE128", "1234567890",
                      "0x000000", "0xFFFFFF", "", "100");
        builder.Writeln();

        // Insert a Code 39 barcode.
        InsertBarcode(builder, "CODE39", "CODE39DATA",
                      "0x000000", "0xFFFFFF", "", "100");
        builder.Writeln();

        // Insert an EAN13 barcode.
        InsertBarcode(builder, "EAN13", "0123456789012",
                      "0x000000", "0xFFFFFF", "", "100");
        builder.Writeln();

        // Update all fields so that the barcodes are rendered in the document.
        doc.UpdateFields();

        // Save the document to a DOCX file.
        doc.Save("Barcodes.docx");
    }
}
