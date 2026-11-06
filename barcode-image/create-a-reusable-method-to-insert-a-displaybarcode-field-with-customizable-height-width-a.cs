using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a barcode with custom height, width and type.
        InsertDisplayBarcode(builder, "1234567890", "CODE128", 50, 200);

        // Update fields to render the barcode.
        doc.UpdateFields();

        // Save the document.
        doc.Save("BarcodeOutput.docx");
    }

    // Reusable method to insert a DISPLAYBARCODE field with customizable parameters.
    private static void InsertDisplayBarcode(DocumentBuilder builder, string value, string type, int height, int width)
    {
        // Insert a typed DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set barcode type and value.
        field.BarcodeType = type;      // e.g., "CODE128", "QR", "EAN13", etc.
        field.BarcodeValue = value;    // Text to encode.

        // Approximate size using the scaling factor (percentage). This is a simple heuristic.
        int scaling = (height + width) / 2;
        field.ScalingFactor = scaling.ToString();

        // Optional: set foreground and background colors.
        field.ForegroundColor = "0x000000"; // Black.
        field.BackgroundColor = "0xFFFFFF"; // White.
    }
}
