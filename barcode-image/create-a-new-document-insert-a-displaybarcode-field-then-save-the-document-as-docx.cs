using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure the barcode (example: QR code with custom colors and scaling).
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "1234567890";
        barcodeField.BackgroundColor = "0xFFFFFF"; // white background
        barcodeField.ForegroundColor = "0x000000"; // black bars
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        // Ensure the field result is updated.
        doc.UpdateFields();

        // Save the document as DOCX.
        doc.Save("DisplayBarcode.docx");
    }
}
