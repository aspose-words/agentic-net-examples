using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a barcode image stream using Aspose.BarCode.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        var stream = new MemoryStream();
        using (var generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            // Set barcode and background colors.
            generator.Parameters.Barcode.BarColor = Color.Black;
            generator.Parameters.BackColor = Color.White;

            // Save as PNG to the stream.
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    // Legacy method delegates to the main implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Expect the first argument to be the input directory path.
        if (args.Length == 0)
            return;

        string inputDir = args[0];
        if (!Directory.Exists(inputDir))
            return;

        // Create an output directory for PDFs.
        string outputDir = Path.Combine(inputDir, "BarcodedPdfs");
        Directory.CreateDirectory(outputDir);

        // Process all .docx files in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the Word document.
            Document doc = new Document(filePath);

            // Register the custom barcode generator for rendered output (PDF).
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Insert a barcode field at the beginning of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentStart();
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            barcodeField.BarcodeType = "CODE128";
            // Use the file name (without extension) as the barcode value.
            barcodeField.BarcodeValue = Path.GetFileNameWithoutExtension(filePath);
            barcodeField.ForegroundColor = "0x000000";
            barcodeField.BackgroundColor = "0xFFFFFF";

            // Update fields to render the barcode.
            doc.UpdateFields();

            // Save the document as PDF in the output directory.
            string outputFileName = Path.GetFileNameWithoutExtension(filePath) + ".pdf";
            string outputPath = Path.Combine(outputDir, outputFileName);
            doc.Save(outputPath);
        }
    }
}
