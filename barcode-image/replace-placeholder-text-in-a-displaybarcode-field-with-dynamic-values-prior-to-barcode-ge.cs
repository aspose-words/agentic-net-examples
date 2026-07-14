using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with a placeholder value.
        // Use the typed insertion method as required.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "CODE39";
        barcodeField.BarcodeValue = "PLACEHOLDER";

        // Dynamic value that should replace the placeholder.
        string dynamicValue = "987654321";

        // Replace the placeholder text with the dynamic value before updating fields.
        if (barcodeField.BarcodeValue == "PLACEHOLDER")
        {
            barcodeField.BarcodeValue = dynamicValue;
        }

        // Update fields so the barcode is generated with the new value.
        doc.UpdateFields();

        // Save the document.
        doc.Save("DisplayBarcodeDynamic.docx");
    }
}
