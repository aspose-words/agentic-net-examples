using System;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field that will render a QR code.
        // The field is inserted with the "true" argument to add a field separator.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set the barcode type and the value to encode.
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "ABC123";

        // Optional: customize appearance – background/foreground colors, error correction, scaling, size, rotation.
        barcodeField.BackgroundColor = "0xF8BD69";
        barcodeField.ForegroundColor = "0xB5413B";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        // Move to the next line after the field.
        builder.Writeln();

        // Save the document in DOCX format.
        doc.Save("BarcodeDocument.docx");
    }
}
