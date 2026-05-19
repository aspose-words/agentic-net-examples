using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a DISPLAYBARCODE field using the typed API.
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Set the barcode properties.
            barcodeField.BarcodeType = "QR";               // Type of barcode (e.g., QR, CODE128, etc.).
            barcodeField.BarcodeValue = "AsposeDemo";      // Data to encode.
            barcodeField.BackgroundColor = "0xFFFFFF";    // Optional: white background.
            barcodeField.ForegroundColor = "0x000000";    // Optional: black foreground.
            barcodeField.ErrorCorrectionLevel = "3";      // Optional: QR error correction level.
            barcodeField.ScalingFactor = "250";           // Optional: scaling factor (percentage).
            barcodeField.SymbolHeight = "1000";           // Optional: height in twips.
            barcodeField.SymbolRotation = "0";            // Optional: rotation (0‑3).

            // Update fields to ensure the field result is generated.
            doc.UpdateFields();

            // Save the document as DOCX.
            doc.Save("DisplayBarcode.docx");
        }
    }
}
