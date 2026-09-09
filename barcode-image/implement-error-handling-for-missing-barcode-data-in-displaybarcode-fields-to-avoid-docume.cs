using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with valid data.
        FieldDisplayBarcode validField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        validField.BarcodeType = "CODE39";
        validField.BarcodeValue = "12345ABCDE";
        validField.AddStartStopChar = true;
        builder.Writeln();

        // Insert a DISPLAYBARCODE field with missing barcode value.
        FieldDisplayBarcode missingField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        missingField.BarcodeType = "CODE39";
        // Intentionally leave BarcodeValue empty to simulate missing data.
        missingField.BarcodeValue = string.Empty;
        missingField.AddStartStopChar = true;
        builder.Writeln();

        // Update fields to generate initial results.
        doc.UpdateFields();

        // Iterate over all fields and handle missing barcode data.
        foreach (Field field in doc.Range.Fields)
        {
            if (field is FieldDisplayBarcode barcodeField)
            {
                // If the barcode value is null or empty, replace it with a placeholder.
                if (string.IsNullOrWhiteSpace(barcodeField.BarcodeValue))
                {
                    // Set a default value that will not cause generation errors.
                    barcodeField.BarcodeValue = "0";
                    // Mark the field as dirty so it will be re‑evaluated.
                    barcodeField.IsDirty = true;
                }
            }
        }

        // Re‑update fields after fixing missing data.
        doc.UpdateFields();

        // Save the document with error handling.
        try
        {
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
            doc.Save(outputPath);
            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error saving document: " + ex.Message);
        }
    }
}
