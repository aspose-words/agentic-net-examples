using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple PNG logo and save it to a static path.
        string logoPath = Path.Combine(outputDir, "logo.png");
        CreateSampleLogo(logoPath);

        // Build a mail‑merge template that contains an image merge field.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        // The field name after "Image:" tells the engine this is an image field.
        builder.InsertField("MERGEFIELD Image:Logo");
        builder.Writeln();
        // Add a regular text merge field for demonstration.
        builder.InsertField("MERGEFIELD Name");
        builder.Writeln();

        // Assign a callback that will supply the static logo image for every image field.
        template.MailMerge.FieldMergingCallback = new StaticImageCallback(logoPath);

        // Create a data source. The value in the "Logo" column is ignored by the callback.
        DataTable data = new DataTable("Employees");
        data.Columns.Add("Logo");
        data.Columns.Add("Name");
        data.Rows.Add("IgnoredValue", "Alice Johnson");
        data.Rows.Add("IgnoredValue", "Bob Smith");

        // Perform the mail merge.
        template.MailMerge.Execute(data);

        // Save the merged document.
        string resultPath = Path.Combine(outputDir, "Merged.docx");
        template.Save(resultPath);
    }

    // Writes a minimal 1x1 PNG image to the specified file.
    private static void CreateSampleLogo(string path)
    {
        // PNG data for a 1x1 transparent pixel.
        byte[] pngBytes = new byte[]
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
        File.WriteAllBytes(path, pngBytes);
    }

    // Callback that supplies the same image for every image merge field.
    private class StaticImageCallback : IFieldMergingCallback
    {
        private readonly string _imagePath;

        public StaticImageCallback(string imagePath)
        {
            _imagePath = imagePath;
        }

        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // No custom handling for text fields.
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Provide the static image file name; the engine will load it.
            args.ImageFileName = _imagePath;
        }
    }
}
