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

        // Insert a DISPLAYBARCODE field.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure the barcode (example: QR code).
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "HelloWorld";
        barcodeField.BackgroundColor = "0xFFFFFF";
        barcodeField.ForegroundColor = "0x000000";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        // Update fields to apply the changes.
        doc.UpdateFields();

        // Save the document as DOCX.
        doc.Save("DisplayBarcode.docx");
    }
}
