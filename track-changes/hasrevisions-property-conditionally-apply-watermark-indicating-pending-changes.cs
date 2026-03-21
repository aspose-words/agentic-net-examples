using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document, or create a new one if the file does not exist.
        Document doc;
        const string inputPath = "Input.docx";
        if (File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            doc = new Document();
            var tmpBuilder = new DocumentBuilder(doc);
            tmpBuilder.Writeln("Sample document created because Input.docx was not found.");
        }

        // If the document has revisions, add a watermark indicating pending changes.
        if (doc.HasRevisions)
        {
            var builder = new DocumentBuilder(doc);
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            var watermark = new Shape(doc, ShapeType.TextPlainText);
            watermark.TextPath.Text = "PENDING CHANGES";
            watermark.TextPath.FontFamily = "Arial";
            watermark.Width = 500;
            watermark.Height = 100;
            watermark.Rotation = -40;
            watermark.Fill.Color = Color.LightGray;
            watermark.StrokeColor = Color.LightGray;

            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.VerticalAlignment = VerticalAlignment.Center;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;

            builder.InsertNode(watermark);
        }

        // Remove all fields (e.g., barcode fields) to avoid runtime errors when saving.
        for (int i = doc.Range.Fields.Count - 1; i >= 0; i--)
        {
            doc.Range.Fields[i].Remove();
        }

        // Save the document as PDF. Disable field updating to avoid processing fields.
        var saveOptions = new PdfSaveOptions
        {
            UpdateFields = false
        };
        doc.Save("Output.pdf", saveOptions);
    }
}
