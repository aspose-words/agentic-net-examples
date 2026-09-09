using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeFieldExample
{
    public class Program
    {
        // Inserts a DISPLAYBARCODE field with the specified parameters.
        // height and scalingFactor are strings because the field properties expect string values.
        public static void InsertDisplayBarcode(DocumentBuilder builder, string value, string type, string height, string scalingFactor)
        {
            // Create a typed DISPLAYBARCODE field.
            var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Set required properties.
            field.BarcodeValue = value;
            field.BarcodeType = type;
            field.SymbolHeight = height;        // Height in TWIPS (1/1440 inch).
            field.ScalingFactor = scalingFactor; // Percentage scaling.

            // Add a line break after the field for readability.
            builder.Writeln();
        }

        public static void Main()
        {
            // Create a new document and builder.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Example 1: EAN13 barcode.
            InsertDisplayBarcode(builder, "501234567890", "EAN13", "1000", "250");

            // Example 2: QR code with larger size.
            InsertDisplayBarcode(builder, "ABC123", "QR", "1500", "300");

            // Example 3: CODE39 barcode with default size.
            InsertDisplayBarcode(builder, "12345ABCDE", "CODE39", "800", "200");

            // Update all fields to ensure they are calculated.
            doc.UpdateFields();

            // Save the document to the current directory.
            doc.Save("DisplayBarcode.docx");
        }
    }
}
