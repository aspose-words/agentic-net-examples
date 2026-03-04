using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeFieldExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a MERGEBARCODE field. The second argument (true) updates the field result immediately.
            FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);

            // Configure the barcode properties.
            barcodeField.BarcodeType = "QR";                 // QR code type.
            barcodeField.BarcodeValue = "ABC123";            // Data to encode.
            barcodeField.BackgroundColor = "0xF8BD69";       // Light orange background.
            barcodeField.ForegroundColor = "0xB5413B";       // Dark red bars.
            barcodeField.ErrorCorrectionLevel = "3";        // Highest error correction.
            barcodeField.ScalingFactor = "250";              // 250 % scaling.
            barcodeField.SymbolHeight = "1000";              // Height in TWIPS.
            barcodeField.SymbolRotation = "0";               // No rotation.

            // Add a paragraph break after the field for readability.
            builder.Writeln();

            // Save the document in DOCX format.
            doc.Save("BarcodeMerge.docx");
        }
    }
}
