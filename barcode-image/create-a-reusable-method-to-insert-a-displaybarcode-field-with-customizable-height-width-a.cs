using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeFieldExample
{
    public class Program
    {
        // Inserts a DISPLAYBARCODE field with the specified parameters.
        // heightTwips – height of the barcode symbol in TWIPS (1/1440 inch).
        // scalingFactor – scaling factor as a percentage (e.g., 250 for 250%).
        // barcodeType – type string understood by Word (e.g., "QR", "CODE39").
        // barcodeValue – data to encode.
        private static void InsertDisplayBarcode(DocumentBuilder builder, string barcodeValue, string barcodeType, int heightTwips, int scalingFactor)
        {
            // Insert a typed DISPLAYBARCODE field.
            var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Set required properties.
            field.BarcodeValue = barcodeValue;
            field.BarcodeType = barcodeType;
            field.SymbolHeight = heightTwips.ToString();      // Height in TWIPS.
            field.ScalingFactor = scalingFactor.ToString();   // Scaling factor as percentage.

            // Update fields so the result is calculated.
            builder.Document.UpdateFields();
        }

        public static void Main()
        {
            // Create a new empty document.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Example usage: insert three different barcodes.
            InsertDisplayBarcode(builder, "ABC123", "QR", 1000, 250);          // QR code, 1000 TWIPS height, 250% scaling.
            builder.Writeln(); // Add a line break.

            InsertDisplayBarcode(builder, "12345ABCDE", "CODE39", 800, 200);   // CODE39 barcode.
            builder.Writeln();

            InsertDisplayBarcode(builder, "501234567890", "EAN13", 1200, 300); // EAN13 barcode with larger size.
            builder.Writeln();

            // Save the document.
            doc.Save("DisplayBarcodes.docx");
        }
    }
}
