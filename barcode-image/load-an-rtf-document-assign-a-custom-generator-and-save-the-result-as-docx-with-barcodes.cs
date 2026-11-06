using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Loading;
using Aspose.BarCode.Generation;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set colors if needed
            generator.Parameters.Barcode.BarColor = Aspose.Drawing.Color.Black;
            generator.Parameters.BackColor = Aspose.Drawing.Color.White;
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
        // Sample RTF content with a simple paragraph.
        string rtfContent = @"{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Helvetica;}}\viewkind4\uc1\pard\fs20 Sample RTF document.\par}";
        // Load the RTF document from a memory stream.
        using (var rtfStream = new MemoryStream(Encoding.UTF8.GetBytes(rtfContent)))
        {
            var loadOptions = new LoadOptions { LoadFormat = LoadFormat.Rtf };
            Document doc = new Document(rtfStream, loadOptions);

            // Insert a barcode field.
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            barcodeField.BarcodeType = "CODE128";
            barcodeField.BarcodeValue = "1234567890";
            barcodeField.ForegroundColor = "0x000000";
            barcodeField.BackgroundColor = "0xFFFFFF";

            // Register the custom barcode generator.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Update fields to render the barcode.
            doc.UpdateFields();

            // Save the document as DOCX.
            doc.Save("output.docx");
        }
    }
}
