using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Inserts a DISPLAYBARCODE field with the specified parameters.
    private static void InsertBarcode(DocumentBuilder builder, string type, string value, string foreColor, string backColor, string errorCorrection = null, string scaling = null)
    {
        // Create a DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = type;               // e.g., "QR", "CODE128"
        field.BarcodeValue = value;             // barcode text
        field.ForegroundColor = foreColor;      // e.g., "0x000000"
        field.BackgroundColor = backColor;      // e.g., "0xFFFFFF"

        // Optional switches.
        if (!string.IsNullOrEmpty(errorCorrection))
            field.ErrorCorrectionLevel = errorCorrection; // for QR codes
        if (!string.IsNullOrEmpty(scaling))
            field.ScalingFactor = scaling;                // percentage scaling
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a QR code.
        InsertBarcode(builder, "QR", "https://example.com", "0x000000", "0xFFFFFF", errorCorrection: "3", scaling: "150");
        builder.Writeln();

        // Insert a CODE128 barcode.
        InsertBarcode(builder, "CODE128", "1234567890", "0x0000FF", "0xFFFFFF", scaling: "120");
        builder.Writeln();

        // Insert a CODE39 barcode.
        InsertBarcode(builder, "CODE39", "CODE39DATA", "0xFF0000", "0xFFFFFF");
        builder.Writeln();

        // Insert an EAN13 barcode.
        InsertBarcode(builder, "EAN13", "5901234123457", "0x000000", "0xFFFFFF", scaling: "100");

        // Update all fields so that the barcode data is stored.
        doc.UpdateFields();

        // Save the document.
        doc.Save("Barcodes.docx");
    }
}
