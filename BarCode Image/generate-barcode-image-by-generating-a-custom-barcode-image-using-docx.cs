using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field that will render a QR code.
        // The field is inserted with a placeholder value; we then set its properties.
        FieldDisplayBarcode barcode = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcode.BarcodeType = "QR";                 // Type of barcode.
        barcode.BarcodeValue = "ABC123";            // Data to encode.
        barcode.BackgroundColor = "0xF8BD69";       // Custom background colour.
        barcode.ForegroundColor = "0xB5413B";       // Custom foreground colour.
        barcode.ErrorCorrectionLevel = "3";         // QR error correction level.
        barcode.ScalingFactor = "250";              // Scale the symbol (percentage).
        barcode.SymbolHeight = "1000";              // Height in TWIPS (1/1440 inch).
        barcode.SymbolRotation = "0";               // No rotation.

        // Add a line break after the barcode for readability.
        builder.Writeln();

        // Save the document in DOCX format.
        doc.Save("CustomBarcode.docx");
    }
}
