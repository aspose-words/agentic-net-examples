using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image merge field. The field name after the "Image:" prefix is "Photo".
        builder.InsertField("MERGEFIELD Image:Photo");

        // Prepare a data source that contains the image data as a byte array.
        DataTable table = new DataTable("Images");
        table.Columns.Add("Photo", typeof(byte[]));

        // A minimal PNG image (1x1 pixel, transparent) encoded in Base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XbZcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Add a single row containing the image bytes.
        DataRow row = table.NewRow();
        row["Photo"] = imageBytes;
        table.Rows.Add(row);

        // Assign a callback that will handle the image field merging.
        doc.MailMerge.FieldMergingCallback = new ImageFieldHandler();

        // Execute the mail merge using the data table.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedImage.docx");
        doc.Save(outputPath);
    }

    // Callback implementation for handling image merge fields.
    private class ImageFieldHandler : IFieldMergingCallback
    {
        // No custom processing for regular text fields.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Intentionally left blank.
        }

        // Called when an image merge field is encountered.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // The field value is expected to be a byte[] containing image data.
            if (args.FieldValue is byte[] imageBytes)
            {
                // Provide the image to the mail merge engine via a stream.
                args.ImageStream = new MemoryStream(imageBytes);
            }
        }
    }
}
