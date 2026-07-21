using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Use Code128 symbology for this example
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Render the barcode to a PNG stream
            var ms = new MemoryStream();
            generator.Save(ms, BarCodeImageFormat.Png);
            ms.Position = 0;
            return ms;
        }
    }

    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        // Legacy behavior mirrors the current implementation
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new empty document
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a typed barcode field
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "Code128";
        field.BarcodeValue = "1234567890";

        // Register the custom barcode generator
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Update fields to render the barcode
        doc.UpdateFields();

        // Save the document as PDF (rendered barcode image will be embedded)
        doc.Save("BarcodeOutput.pdf", SaveFormat.Pdf);
    }
}
