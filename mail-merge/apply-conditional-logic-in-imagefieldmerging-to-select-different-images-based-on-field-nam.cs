using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some text.
        builder.Writeln("Mail merge with conditional image selection:");

        // Insert image merge fields.
        builder.InsertField("MERGEFIELD Logo");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Signature");

        // Set up the image field merging callback.
        doc.MailMerge.FieldMergingCallback = new ImageFieldMergingHandler();

        // Execute mail merge with placeholder values (null) for the image fields.
        string[] fieldNames = { "Logo", "Signature" };
        object[] fieldValues = { null, null };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        string outputPath = "Merged.docx";
        doc.Save(outputPath);

        // Validate that images were inserted.
        int imageCount = doc.GetChildNodes(NodeType.Shape, true).Count;
        if (imageCount == 0)
            throw new Exception("No images were merged into the document.");

        Console.WriteLine($"Document '{outputPath}' created with {imageCount} images.");
    }

    // Callback class to provide images based on field name.
    private class ImageFieldMergingHandler : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            // No text field handling required for this example.
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Choose image bytes based on the field name.
            byte[] imageBytes = GetImageBytes(args.FieldName);
            if (imageBytes == null || imageBytes.Length == 0)
                throw new Exception($"No image data defined for field '{args.FieldName}'.");

            // Provide the image stream to the merge engine.
            MemoryStream imageStream = new MemoryStream(imageBytes);
            args.ImageStream = imageStream;
            args.ImageFileName = $"{args.FieldName}.png";
        }

        private static byte[] GetImageBytes(string fieldName)
        {
            // Simple 1x1 PNG images encoded in base64.
            // Red pixel for "Logo", green pixel for "Signature", default to blue.
            const string redBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lKXcAAAAAElFTkSuQmCC";
            const string greenBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8DAAQAB/6F9WQAAAABJRU5ErkJggg==";
            const string blueBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8DAAQAB/6F9WQAAAABJRU5ErkJggg==";

            string base64 = fieldName.Equals("Logo", StringComparison.OrdinalIgnoreCase)
                ? redBase64
                : fieldName.Equals("Signature", StringComparison.OrdinalIgnoreCase)
                    ? greenBase64
                    : blueBase64;

            return Convert.FromBase64String(base64);
        }
    }
}
