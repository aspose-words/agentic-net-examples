using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeDemo
{
    public class Program
    {
        // Inserts a DISPLAYBARCODE field with the specified parameters.
        // heightInTwips and scalingFactor are optional; pass null to leave default.
        private static void InsertDisplayBarcode(DocumentBuilder builder, string barcodeValue, string barcodeType, string heightInTwips = null, string scalingFactor = null)
        {
            // Create a typed DISPLAYBARCODE field.
            var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Set required properties.
            field.BarcodeValue = barcodeValue;
            field.BarcodeType = barcodeType;

            // Optional customizations.
            if (!string.IsNullOrEmpty(heightInTwips))
                field.SymbolHeight = heightInTwips;          // Height in TWIPS (1/1440 inch).

            if (!string.IsNullOrEmpty(scalingFactor))
                field.ScalingFactor = scalingFactor;         // Scaling factor as a percentage (e.g., "250").

            // Update the document fields so the result reflects the new settings.
            builder.Document.UpdateFields();
        }

        public static void Main()
        {
            // Create a new blank document.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Example 1: EAN13 barcode with custom height and scaling.
            InsertDisplayBarcode(builder, "501234567890", "EAN13", "1000", "250");
            builder.Writeln();

            // Example 2: QR code with different size settings.
            InsertDisplayBarcode(builder, "ABC123", "QR", "1200", "300");
            builder.Writeln();

            // Example 3: CODE39 barcode with default size.
            InsertDisplayBarcode(builder, "12345ABCDE", "CODE39");
            builder.Writeln();

            // Save the document to the current directory.
            doc.Save("DisplayBarcode.docx");
        }
    }
}
