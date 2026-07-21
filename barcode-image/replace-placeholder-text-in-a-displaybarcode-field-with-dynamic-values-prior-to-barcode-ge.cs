using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and insert a DISPLAYBARCODE field with a placeholder value.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the field using the typed API.
        FieldDisplayBarcode placeholderField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        placeholderField.BarcodeType = "CODE39";
        placeholderField.BarcodeValue = "PLACEHOLDER"; // placeholder text
        placeholderField.AddStartStopChar = true;      // optional appearance setting
        builder.Writeln();

        // Save the document that contains the placeholder field.
        doc.Save("Placeholder.docx");

        // Load the document again to simulate a later processing step.
        Document loadedDoc = new Document("Placeholder.docx");

        // Locate the DISPLAYBARCODE field and replace the placeholder with the actual value.
        foreach (Field field in loadedDoc.Range.Fields)
        {
            if (field.Type == FieldType.FieldDisplayBarcode)
            {
                FieldDisplayBarcode displayBarcode = (FieldDisplayBarcode)field;
                if (displayBarcode.BarcodeValue == "PLACEHOLDER")
                {
                    displayBarcode.BarcodeValue = "12345ABCDE"; // dynamic value
                }
            }
        }

        // Update fields so the barcode is generated with the new value.
        loadedDoc.UpdateFields();

        // Save the final document.
        loadedDoc.Save("Result.docx");
    }
}
