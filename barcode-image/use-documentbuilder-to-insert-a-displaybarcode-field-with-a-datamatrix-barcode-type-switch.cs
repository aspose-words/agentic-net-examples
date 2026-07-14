using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace DisplayBarcodeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a DISPLAYBARCODE field using the typed API.
            // The second argument (true) tells the builder to insert the field as a separate node.
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Set the barcode type to DataMatrix and provide a sample value.
            barcodeField.BarcodeType = "DataMatrix";
            barcodeField.BarcodeValue = "1234567890";

            // Optionally, customize appearance (colors, scaling, etc.).
            // barcodeField.BackgroundColor = "0xFFFFFF";
            // barcodeField.ForegroundColor = "0x000000";

            // Update fields to ensure the field result is generated.
            doc.UpdateFields();

            // Save the document to the local file system.
            doc.Save("DisplayBarcodeDataMatrix.docx");
        }
    }
}
