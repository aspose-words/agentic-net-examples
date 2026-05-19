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
        // Create a new blank document.
        Document doc = new Document();

        // Insert a MERGEFIELD that will accept an image during mail merge.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertField("MERGEFIELD Image:Photo");

        // Prepare a simple data source with a single column that matches the merge field name.
        DataTable table = new DataTable("Images");
        table.Columns.Add("Photo");
        table.Rows.Add("placeholder"); // Value is not used because we supply the image in the callback.

        // Assign a callback that will provide the image and adjust its size.
        doc.MailMerge.FieldMergingCallback = new ImageResizerCallback();

        // Execute the mail merge.
        doc.MailMerge.Execute(table);

        // Update fields to reflect the merged content.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("MergedImage.docx");
    }

    // Callback that supplies an image and modifies its size during mail merge.
    private class ImageResizerCallback : IFieldMergingCallback
    {
        // No custom processing required for text fields.
        public void FieldMerging(FieldMergingArgs args)
        {
        }

        // Called when an image merge field is encountered.
        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // A tiny 1x1 pixel PNG (transparent) encoded in Base64.
            const string base64Png =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
                "Z6ZcAAAAASUVORK5CYII=";

            // Decode the Base64 string to a byte array.
            byte[] imageBytes = Convert.FromBase64String(base64Png);

            // Write the image to a temporary file.
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
            File.WriteAllBytes(tempPath, imageBytes);

            // Supply the image file name to the mail merge engine.
            args.ImageFileName = tempPath;

            // Adjust the image size to 50 % of its original dimensions.
            // Use MergeFieldImageDimension with Percent unit for both width and height.
            args.ImageWidth = new MergeFieldImageDimension(0.5, MergeFieldImageDimensionUnit.Percent);
            args.ImageHeight = new MergeFieldImageDimension(0.5, MergeFieldImageDimensionUnit.Percent);
        }
    }
}
