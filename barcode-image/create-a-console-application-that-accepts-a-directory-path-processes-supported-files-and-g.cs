using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.BarCode.Generation;

public sealed class CustomBarcodeGenerator : IBarcodeGenerator
{
    // Generates a PNG barcode image using Aspose.BarCode and returns it as a stream.
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

    // Legacy method – delegates to the primary implementation.
    public Stream GetOldBarcodeImage(Aspose.Words.Fields.BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Determine the directory to process (first argument or current directory).
        string inputDir = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();
        if (!Directory.Exists(inputDir))
            return;

        // Create an output subfolder for the generated PDFs.
        string outputDir = Path.Combine(inputDir, "output");
        Directory.CreateDirectory(outputDir);

        // Process each .docx file in the input directory.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the Word document.
            Document doc = new Document(filePath);

            // Insert a barcode field at the beginning of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentStart();
            FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
            field.BarcodeType = "CODE128";
            field.BarcodeValue = Path.GetFileNameWithoutExtension(filePath);
            field.ForegroundColor = "0x000000";
            field.BackgroundColor = "0xFFFFFF";

            // Register the custom barcode generator for rendered output (PDF).
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Update fields so the barcode is rendered.
            doc.UpdateFields();

            // Save the document as PDF in the output folder.
            string pdfPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }
    }
}
