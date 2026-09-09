using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with a placeholder value.
        Aspose.Words.Fields.FieldDisplayBarcode displayBarcode = 
            (Aspose.Words.Fields.FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        displayBarcode.BarcodeType = "QR";
        displayBarcode.BarcodeValue = "PLACEHOLDER"; // placeholder text
        displayBarcode.BackgroundColor = "0xF8BD69";
        displayBarcode.ForegroundColor = "0xB5413B";
        displayBarcode.ErrorCorrectionLevel = "3";
        displayBarcode.ScalingFactor = "250";
        displayBarcode.SymbolHeight = "1000";
        displayBarcode.SymbolRotation = "0";

        // Add a line break after the field.
        builder.Writeln();

        // Replace the placeholder with a dynamic value before generating the barcode.
        // In a real scenario this value could come from a database, user input, etc.
        string dynamicValue = "ABC123";
        displayBarcode.BarcodeValue = dynamicValue;

        // Update all fields in the document so the barcode image is generated.
        doc.UpdateFields();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DisplayBarcodeDynamic.docx");
        doc.Save(outputPath);
    }
}
