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

        // Insert a MERGEBARCODE field that will generate a QR code.
        FieldMergeBarcode mergeBarcode = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeBarcode.BarcodeType = "QR";                     // Set barcode type.
        mergeBarcode.BarcodeValue = "MyQRCode";              // Data to encode.
        mergeBarcode.BackgroundColor = "F8BD69";            // Background color (hex string, without 0x).
        mergeBarcode.ForegroundColor = "B5413B";            // Foreground color (hex string, without 0x).
        mergeBarcode.ErrorCorrectionLevel = "3";            // QR error correction level.
        mergeBarcode.ScalingFactor = "250";                 // Scale the symbol.
        mergeBarcode.SymbolHeight = "1000";                // Height in TWIPS.
        mergeBarcode.SymbolRotation = "0";                 // No rotation.

        // Add a paragraph break after the field.
        builder.Writeln();

        // Save the document to a DOCX file.
        doc.Save("BarcodeQR.docx");
    }
}
