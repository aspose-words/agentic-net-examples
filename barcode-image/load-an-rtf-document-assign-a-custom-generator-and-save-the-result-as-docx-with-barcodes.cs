using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Optional: set colors if needed
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;
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
        // Prepare a simple RTF file to load
        string tempRtfPath = Path.Combine(Path.GetTempPath(), "input.rtf");
        File.WriteAllText(tempRtfPath, @"{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Helvetica;}}\f0\fs24 Sample RTF document.}");

        // Load the RTF document
        Document doc = new Document(tempRtfPath);

        // Register the custom barcode generator
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Insert a barcode field using typed construction
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "CODE128";
        barcodeField.BarcodeValue = "1234567890";
        barcodeField.ForegroundColor = "0x000000";
        barcodeField.BackgroundColor = "0xFFFFFF";

        // Update fields to render the barcode
        doc.UpdateFields();

        // Save as DOCX
        string outputPath = Path.Combine(Path.GetTempPath(), "output.docx");
        doc.Save(outputPath);
    }
}
