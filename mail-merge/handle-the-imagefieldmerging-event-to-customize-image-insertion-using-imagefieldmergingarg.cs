using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a temporary directory for the sample files.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeWordsSample");
        Directory.CreateDirectory(tempDir);

        // Create a simple 1x1 red PNG image and save it to a file.
        // The PNG byte array represents a minimal red pixel image.
        byte[] pngData = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x02,0x00,0x00,0x00,0x90,0x77,0x53,
            0xDE,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x08,0xD7,0x63,0x60,0x60,0x60,0x00,
            0x00,0x00,0x05,0x00,0x01,0x0D,0x0A,0x2D,
            0xB4,0x00,0x00,0x00,0x00,0x49,0x45,0x4E,
            0x44,0xAE,0x42,0x60,0x82
        };
        string imagePath = Path.Combine(tempDir, "sample.png");
        File.WriteAllBytes(imagePath, pngData);

        // Build a source document that contains an image merge field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The field name is prefixed with "Image:" so the mail merge engine treats it as an image field.
        builder.InsertField("MERGEFIELD Image:ImageColumn");

        // Prepare a data source with a single column that holds the image file name.
        DataTable table = new DataTable("Images");
        table.Columns.Add("ImageColumn", typeof(string));
        table.Rows.Add(imagePath); // Use the path of the image we just created.

        // Assign a callback that will supply the image during the merge.
        doc.MailMerge.FieldMergingCallback = new ImageFieldHandler();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        string outputPath = Path.Combine(tempDir, "MergedDocument.docx");
        doc.Save(outputPath);

        // Optional: clean up the temporary image file.
        // File.Delete(imagePath);
    }

    // Callback implementation that handles image merge fields.
    private class ImageFieldHandler : IFieldMergingCallback
    {
        // No custom handling for regular text fields.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // Intentionally left blank.
        }

        // Called when an image merge field is encountered.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // The field value contains the file name of the image.
            string fileName = args.FieldValue.ToString();

            // Use the ImageFileName property to let Aspose.Words load the image.
            args.ImageFileName = fileName;
        }
    }
}
