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

        // Set required properties for the barcode.
        barcodeField.BarcodeType = "QR";               // Type of barcode.
        barcodeField.BarcodeValue = "ABC123";          // Data to encode.
        barcodeField.BackgroundColor = "0xF8BD69";    // Optional background color.
        barcodeField.ForegroundColor = "0xB5413B";    // Optional foreground color.
        barcodeField.ErrorCorrectionLevel = "3";      // QR specific parameter.
        barcodeField.ScalingFactor = "250";           // Scale the symbol.
        barcodeField.SymbolHeight = "1000";           // Height in twips.
        barcodeField.SymbolRotation = "0";            // Rotation.

        // Update fields to ensure the barcode result is generated.
        doc.UpdateFields();

        // Save the document as DOCX.
        doc.Save("DisplayBarcode.docx");
    }
}
