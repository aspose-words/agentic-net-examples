using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with missing barcode value.
        FieldDisplayBarcode field1 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field1.BarcodeType = "CODE128";
        field1.BarcodeValue = ""; // Missing data.
        field1.ForegroundColor = "0x000000";
        field1.BackgroundColor = "0xFFFFFF";

        // Insert another DISPLAYBARCODE field with proper data.
        FieldDisplayBarcode field2 = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field2.BarcodeType = "QR";
        field2.BarcodeValue = "SampleData";
        field2.ForegroundColor = "0x000000";
        field2.BackgroundColor = "0xFFFFFF";

        // Register custom barcode generator for rendered output (PDF).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Validate barcode fields and provide fallback values for missing data.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type == FieldType.FieldDisplayBarcode)
            {
                var barcodeField = (FieldDisplayBarcode)field;
                if (string.IsNullOrWhiteSpace(barcodeField.BarcodeValue))
                {
                    barcodeField.BarcodeValue = "N/A";
                    if (string.IsNullOrWhiteSpace(barcodeField.BarcodeType))
                    {
                        barcodeField.BarcodeType = "CODE128";
                    }
                }
            }
        }

        // Update fields to render barcodes.
        doc.UpdateFields();

        // Save the document as PDF.
        try
        {
            doc.Save("output.pdf");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error saving document: " + ex.Message);
        }
    }
}
