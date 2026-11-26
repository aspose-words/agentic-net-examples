using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image based on the provided parameters.
    public Stream GetBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        MemoryStream stream = new MemoryStream();
        using (BarcodeGenerator generator = new BarcodeGenerator(EncodeTypes.Code128, parameters.BarcodeValue))
        {
            generator.Save(stream, BarCodeImageFormat.Png);
        }
        stream.Position = 0;
        return stream;
    }

    // Legacy method delegates to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main()
    {
        // Define input and output directories (relative to the executable location).
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputPdfs");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // If the input folder does not exist, exit gracefully.
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Process each .doc file in the input folder.
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.doc"))
        {
            // Load the Word document.
            Document doc = new Document(filePath);

            // Insert a barcode field at the beginning of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentStart();
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            barcodeField.BarcodeType = "CODE128";
            barcodeField.BarcodeValue = Path.GetFileNameWithoutExtension(filePath);
            barcodeField.ForegroundColor = "0x000000";
            barcodeField.BackgroundColor = "0xFFFFFF";

            // Register the custom barcode generator for rendered output (PDF).
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Render the barcode by updating fields.
            doc.UpdateFields();

            // Save the document as PDF.
            string outputFile = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
            doc.Save(outputFile);
        }
    }
}
