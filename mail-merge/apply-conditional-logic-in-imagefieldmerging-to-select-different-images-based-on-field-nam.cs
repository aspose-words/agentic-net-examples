using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add explanatory text and two image merge fields.
        builder.Writeln("Below are images inserted via mail merge:");
        builder.InsertField("MERGEFIELD Image1 \\* MERGEFORMAT", null);
        builder.Writeln();
        builder.InsertField("MERGEFIELD Image2 \\* MERGEFORMAT", null);
        builder.Writeln();

        // Assign the custom image field merging callback.
        doc.MailMerge.FieldMergingCallback = new ImageFieldMergingHandler();

        // Execute mail merge. The actual values are not used; the callback supplies the images.
        doc.MailMerge.Execute(
            new[] { "Image1", "Image2" },
            new object[] { null, null });

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}

// Callback that supplies different images based on the merge field name.
public class ImageFieldMergingHandler : IFieldMergingCallback
{
    // Not used for text fields, but must be implemented.
    public void FieldMerging(FieldMergingArgs args)
    {
        // No action needed for text fields in this example.
    }

    // Called for each image merge field.
    public void ImageFieldMerging(ImageFieldMergingArgs args)
    {
        // Choose an image based on the field name and provide it via a stream.
        if (args.FieldName == "Image1")
        {
            args.ImageStream = new MemoryStream(GetRedPixelPng());
        }
        else if (args.FieldName == "Image2")
        {
            args.ImageStream = new MemoryStream(GetGreenPixelPng());
        }

        // Optional: set a file name for the image.
        args.ImageFileName = "image.png";
    }

    // Returns a 1x1 red PNG image as a byte array.
    private static byte[] GetRedPixelPng()
    {
        const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+hHgAFgwJ/lcKZVwAAAABJRU5ErkJggg==";
        return Convert.FromBase64String(base64);
    }

    // Returns a 1x1 green PNG image as a byte array.
    private static byte[] GetGreenPixelPng()
    {
        const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8AAAgEB/6V6WQAAAABJRU5ErkJggg==";
        return Convert.FromBase64String(base64);
    }
}
