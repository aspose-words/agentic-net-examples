using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed API.
        // The field is inserted as a placeholder; we will set its properties next.
        Aspose.Words.Fields.FieldDisplayBarcode barcodeField = 
            (Aspose.Words.Fields.FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure the barcode properties.
        barcodeField.BarcodeType = "QR";               // Type of barcode (e.g., QR, CODE39, EAN13, etc.).
        barcodeField.BarcodeValue = "ABC123";          // Data to encode.
        barcodeField.BackgroundColor = "0xF8BD69";    // Optional background color.
        barcodeField.ForegroundColor = "0xB5413B";    // Optional foreground color.
        barcodeField.ErrorCorrectionLevel = "3";      // QR-specific error correction level.
        barcodeField.ScalingFactor = "250";           // Scale the barcode.
        barcodeField.SymbolHeight = "1000";           // Height in twips.
        barcodeField.SymbolRotation = "0";            // Rotation (0‑3).

        // Update fields to ensure the field result is generated.
        doc.UpdateFields();

        // Save the document as DOCX.
        doc.Save("DisplayBarcode.docx");
    }
}
