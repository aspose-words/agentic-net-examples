// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class SimpleBarcodeGenerator : IBarcodeGenerator
{
    // Generates a simple placeholder barcode image.
    public Image GetBarcodeImage(BarcodeParameters parameters)
    {
        // Define image size (pixels).
        const int width = 300;
        const int height = 100;

        // Create a bitmap and draw the barcode value as text.
        Bitmap bitmap = new Bitmap(width, height);
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            using (Font font = new Font("Arial", 24, FontStyle.Bold, GraphicsUnit.Point))
            {
                string value = parameters?.BarcodeValue ?? "000000";
                graphics.DrawString(value, font, Brushes.Black, new PointF(10, 30));
            }
        }

        // Set the desired DPI (dots per inch) for the image.
        const float dpi = 300f; // example DPI
        bitmap.SetResolution(dpi, dpi);

        return bitmap;
    }

    // For compatibility; reuse the same implementation.
    public Image GetOldBarcodeImage(BarcodeParameters parameters)
    {
        return GetBarcodeImage(parameters);
    }
}

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Assign the custom barcode generator to the document's field options.
        doc.FieldOptions.BarcodeGenerator = new SimpleBarcodeGenerator();

        // Use DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define barcode parameters (type and value can be adjusted as needed).
        BarcodeParameters barcodeParams = new BarcodeParameters
        {
            BarcodeType = "CODE39",
            BarcodeValue = "12345ABCDE"
        };

        // Generate the barcode image using the custom generator.
        Image barcodeImage = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParams);

        // Insert the generated image into the document.
        builder.InsertImage(barcodeImage);

        // Save the document as DOCX.
        doc.Save("BarcodeDocument.docx", SaveFormat.Docx);
    }
}
