using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary PNG image that will be used as the company logo.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "logo.png");
        CreateSampleLogo(imagePath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text merge fields.
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln();

        // Insert an image merge field. The part after "Image:" is the field name used in the data source.
        builder.InsertField("MERGEFIELD Image:CompanyLogo");

        // Build a simple data source. The value for the image field is ignored because the callback supplies a static image.
        DataTable data = new DataTable("Employees");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("CompanyLogo"); // placeholder column
        data.Rows.Add("John", "Doe", "");
        data.Rows.Add("Jane", "Smith", "");

        // Assign a callback that will provide the static logo image for every merge.
        doc.MailMerge.FieldMergingCallback = new StaticImageCallback(imagePath);

        // Perform the mail merge.
        doc.MailMerge.Execute(data);

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.docx");
        doc.Save(outputPath);

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }

    // Writes a minimal 1x1 PNG image to the specified path.
    private static void CreateSampleLogo(string path)
    {
        // This is a base64‑encoded 1×1 pixel transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }

    // Callback that supplies a static image for the "CompanyLogo" merge field.
    private class StaticImageCallback : IFieldMergingCallback
    {
        private readonly string _imagePath;

        public StaticImageCallback(string imagePath)
        {
            _imagePath = imagePath;
        }

        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // No custom text handling required.
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Ensure we only handle the intended merge field.
            if (!string.IsNullOrEmpty(args.FieldName) &&
                args.FieldName.Equals("CompanyLogo", StringComparison.OrdinalIgnoreCase))
            {
                // Use the file name instead of a System.Drawing.Image instance.
                args.ImageFileName = _imagePath;
            }
        }
    }
}
