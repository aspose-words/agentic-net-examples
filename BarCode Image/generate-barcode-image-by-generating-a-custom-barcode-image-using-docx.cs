using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field that will render a QR code.
        // The field is inserted with the "true" argument so that the field code is visible in the document.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure the barcode properties.
        barcodeField.BarcodeType = "QR";                 // QR code type.
        barcodeField.BarcodeValue = "ABC123";            // Data to encode.
        barcodeField.BackgroundColor = "0xF8BD69";       // Custom background colour.
        barcodeField.ForegroundColor = "0xB5413B";       // Custom foreground colour.
        barcodeField.ErrorCorrectionLevel = "3";         // Highest error correction.
        barcodeField.ScalingFactor = "250";              // 250 % scaling.
        barcodeField.SymbolHeight = "1000";              // Height in TWIPS.
        barcodeField.SymbolRotation = "0";               // No rotation.

        // Add a line break after the field for readability.
        builder.Writeln();

        // Save the document in DOCX format.
        doc.Save("CustomBarcode.docx");
    }
}
