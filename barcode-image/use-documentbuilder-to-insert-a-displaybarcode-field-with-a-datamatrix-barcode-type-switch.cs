using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed API.
        // The field is created via InsertField with FieldType.FieldDisplayBarcode.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set the barcode type to DataMatrix and provide a sample value.
        barcodeField.BarcodeType = "DATAMATRIX";
        barcodeField.BarcodeValue = "123456789";

        // Optionally, customize appearance (e.g., scaling factor, colors).
        // barcodeField.ScalingFactor = "250";
        // barcodeField.BackgroundColor = "0xFFFFFF";
        // barcodeField.ForegroundColor = "0x000000";

        // Update fields to ensure the field result is generated.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("DataMatrixBarcode.docx");
    }
}
