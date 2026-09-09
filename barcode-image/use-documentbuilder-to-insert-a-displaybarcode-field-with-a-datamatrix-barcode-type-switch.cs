using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a DISPLAYBARCODE field.
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

            // Set the barcode type to DataMatrix and provide a value.
            barcodeField.BarcodeType = "DataMatrix";
            barcodeField.BarcodeValue = "1234567890";

            // Optional: set background and foreground colors.
            barcodeField.BackgroundColor = "0xFFFFFF";
            barcodeField.ForegroundColor = "0x000000";

            // Update fields to ensure the field result is generated.
            doc.UpdateFields();

            // Save the document to disk.
            doc.Save("DisplayBarcode_DataMatrix.docx");
        }
    }
}
