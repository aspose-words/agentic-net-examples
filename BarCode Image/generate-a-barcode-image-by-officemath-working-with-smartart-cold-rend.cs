using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will render a QR code.
        // The field is inserted and updated immediately.
        FieldMergeBarcode barcode = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        barcode.BarcodeType = "QR";                 // QR code type.
        barcode.BarcodeValue = "ABC123";            // Data encoded in the barcode.
        barcode.BackgroundColor = "0xF8BD69";       // Custom background color.
        barcode.ForegroundColor = "0xB5413B";       // Custom foreground color.
        barcode.ErrorCorrectionLevel = "3";         // Highest error correction.
        barcode.ScalingFactor = "250";              // Scale the symbol.
        barcode.SymbolHeight = "1000";              // Height in TWIPS.
        barcode.SymbolRotation = "0";               // No rotation.

        // Add a line break after the barcode for readability.
        builder.Writeln();

        // Add a text watermark to the document.
        doc.Watermark.SetText("Sample Watermark");

        // Save the document as DOCX.
        doc.Save("BarcodeWithWatermark.docx");
    }
}
