using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the sample image.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeMailMergeSample");
        Directory.CreateDirectory(tempFolder);

        // Write a minimal 1x1 PNG image to a file (no System.Drawing dependency).
        string imagePath = Path.Combine(tempFolder, "sample.png");
        byte[] pngData = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
        File.WriteAllBytes(imagePath, pngData);

        // Build a simple document with an image merge field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The field name is prefixed with "Image:" to indicate an image merge field.
        builder.InsertField("MERGEFIELD Image:Photo");

        // Prepare a data source that contains the file name of the image.
        DataTable table = new DataTable("Images");
        table.Columns.Add("Photo");
        table.Rows.Add(imagePath);

        // Assign a callback that will load the image from the file name.
        doc.MailMerge.FieldMergingCallback = new ImageFilenameCallback();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Save the result.
        string outputPath = Path.Combine(tempFolder, "Result.docx");
        doc.Save(outputPath);
    }

    // Callback that handles image merge fields.
    private class ImageFilenameCallback : IFieldMergingCallback
    {
        // No custom handling for text fields.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Intentionally left blank.
        }

        // Called when an image merge field is encountered.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // The field value contains the file name of the image.
            string fileName = args.FieldValue?.ToString();
            if (!string.IsNullOrEmpty(fileName) && File.Exists(fileName))
            {
                // Use the file name property to supply the image to the mail merge engine.
                args.ImageFileName = fileName;
            }
        }
    }
}
