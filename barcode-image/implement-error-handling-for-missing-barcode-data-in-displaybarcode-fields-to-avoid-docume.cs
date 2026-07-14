using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with a valid value.
        Aspose.Words.Fields.FieldDisplayBarcode validField = (Aspose.Words.Fields.FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        validField.BarcodeType = "QR";
        validField.BarcodeValue = "VALID123";

        // Insert a DISPLAYBARCODE field with missing barcode data (empty string).
        Aspose.Words.Fields.FieldDisplayBarcode missingField = (Aspose.Words.Fields.FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        missingField.BarcodeType = "QR";
        missingField.BarcodeValue = ""; // Intentionally empty to simulate missing data.

        // Ensure the document updates fields before saving.
        doc.UpdateFields();

        // Error handling: replace missing barcode values with a placeholder.
        foreach (Field field in doc.Range.Fields)
        {
            if (field is Aspose.Words.Fields.FieldDisplayBarcode barcodeField)
            {
                if (string.IsNullOrWhiteSpace(barcodeField.BarcodeValue))
                {
                    // Assign a default value to prevent save failures.
                    barcodeField.BarcodeValue = "PLACEHOLDER";
                }
            }
        }

        // Attempt to save the document, catching any unexpected exceptions.
        try
        {
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Barcodes.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving document: {ex.Message}");
        }
    }
}
