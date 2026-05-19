using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Path for the placeholder logo image.
        string logoPath = Path.Combine(Directory.GetCurrentDirectory(), "logo.png");
        CreatePlaceholderLogo(logoPath);

        // Create a simple source document with an image merge field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // The field name after "Image:" must match the column name in the data source.
        builder.InsertField("MERGEFIELD Image:Logo");

        // Build a data source. The actual column value is irrelevant because the callback supplies the image.
        DataTable data = new DataTable("Employees");
        data.Columns.Add("Logo");
        data.Rows.Add("Logo");
        data.Rows.Add("Logo");
        data.Rows.Add("Logo");

        // Assign a callback that provides the same static logo for every merge field.
        doc.MailMerge.FieldMergingCallback = new StaticImageCallback(logoPath);

        // Perform the mail merge.
        doc.MailMerge.Execute(data);

        // Save the merged document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.docx");
        doc.Save(outputPath);
    }

    // Writes a minimal PNG (1x1 pixel) to the specified path.
    private static void CreatePlaceholderLogo(string path)
    {
        // This is a base64‑encoded 1×1 pixel transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, pngBytes);
    }

    // Callback that supplies the same image for every MERGEFIELD with the "Image:" prefix.
    private class StaticImageCallback : IFieldMergingCallback
    {
        private readonly string _imagePath;

        public StaticImageCallback(string imagePath)
        {
            _imagePath = imagePath;
        }

        // No custom text handling required.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args) { }

        // Provide the static image by setting the ImageFileName property.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            args.ImageFileName = _imagePath;
        }
    }
}
