using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.BarCode.Generation;
using Aspose.Drawing;

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
        // Create a new document and insert a barcode field.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";

        // Register the custom barcode generator for rendering.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();
        doc.UpdateFields();

        // Render the document to PDF in memory.
        using (var pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, SaveFormat.Pdf);
            pdfStream.Position = 0;

            // Load the rendered PDF to extract images.
            var pdfDoc = new Document(pdfStream);
            Shape barcodeShape = null;
            foreach (Shape shape in pdfDoc.GetChildNodes(NodeType.Shape, true))
            {
                if (shape.HasImage)
                {
                    barcodeShape = shape;
                    break;
                }
            }

            if (barcodeShape == null)
            {
                Console.WriteLine("No barcode image found in the PDF.");
                return;
            }

            // Get dimensions of the extracted barcode image.
            int extractedWidth, extractedHeight;
            using (var extractedStream = new MemoryStream(barcodeShape.ImageData.ImageBytes))
            using (var extractedImage = Image.FromStream(extractedStream))
            {
                extractedWidth = extractedImage.Width;
                extractedHeight = extractedImage.Height;
            }

            // Generate the barcode image directly using the same generator.
            int generatedWidth, generatedHeight;
            using (var generator = new BarcodeGenerator(EncodeTypes.Code128, "1234567890"))
            using (var genStream = new MemoryStream())
            {
                generator.Save(genStream, BarCodeImageFormat.Png);
                genStream.Position = 0;
                using (var genImage = Image.FromStream(genStream))
                {
                    generatedWidth = genImage.Width;
                    generatedHeight = genImage.Height;
                }
            }

            // Compare dimensions.
            bool dimensionsMatch = extractedWidth == generatedWidth && extractedHeight == generatedHeight;
            Console.WriteLine($"Extracted barcode size: {extractedWidth}x{extractedHeight}");
            Console.WriteLine($"Generated barcode size: {generatedWidth}x{generatedHeight}");
            Console.WriteLine($"Dimensions match: {dimensionsMatch}");
        }
    }
}
