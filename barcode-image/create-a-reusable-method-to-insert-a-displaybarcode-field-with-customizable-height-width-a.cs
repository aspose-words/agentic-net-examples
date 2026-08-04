using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeFieldExample
{
    public class Program
    {
        // Inserts a DISPLAYBARCODE field with the specified parameters.
        // heightTwips – height of the barcode symbol in TWIPS (1/1440 inch).
        // scalingFactor – scaling factor for the symbol (percentage, e.g., "250").
        private static void InsertDisplayBarcode(DocumentBuilder builder, string value, string type, string heightTwips, string scalingFactor)
        {
            // Insert a typed DISPLAYBARCODE field.
            var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Set required properties.
            field.BarcodeValue = value;
            field.BarcodeType = type;
            field.SymbolHeight = heightTwips;
            field.ScalingFactor = scalingFactor;

            // Move to the next line after the field.
            builder.Writeln();
        }

        public static void Main()
        {
            // Create a new empty document.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Example usages of the reusable method.
            InsertDisplayBarcode(builder, "ABC123", "QR", "1000", "250");          // QR code.
            InsertDisplayBarcode(builder, "501234567890", "EAN13", "800", "200"); // EAN13 code.
            InsertDisplayBarcode(builder, "12345ABCDE", "CODE39", "900", "150"); // CODE39 code.
            InsertDisplayBarcode(builder, "09312345678907", "ITF14", "1100", "300"); // ITF14 code.

            // Update fields to ensure the results are calculated.
            doc.UpdateFields();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Barcodes.docx");
            doc.Save(outputPath);
        }
    }
}
