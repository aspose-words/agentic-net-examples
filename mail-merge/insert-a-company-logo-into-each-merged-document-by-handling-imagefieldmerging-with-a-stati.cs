using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeMailMergeExample
{
    // Callback that supplies the static logo image for every image merge field.
    public class StaticLogoCallback : IFieldMergingCallback
    {
        private readonly string _logoPath;

        public StaticLogoCallback(string logoPath)
        {
            _logoPath = logoPath;
        }

        // No custom text merging needed.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args) { }

        // Called for MERGEFIELDs with the "Image:" prefix.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Use the static logo file for every image field.
            args.ImageFileName = _logoPath;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a working directory.
            string workDir = Path.Combine(Path.GetTempPath(), "AsposeMailMergeExample");
            Directory.CreateDirectory(workDir);

            // Create a tiny PNG image (1x1 pixel) from a base‑64 string.
            string logoPath = Path.Combine(workDir, "logo.png");
            const string base64Png =
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BFwAE/wJ/lKXKAAAAAElFTkSuQmCC";
            byte[] logoBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(logoPath, logoBytes);

            // Build a simple mail‑merge template with an image field.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            // Insert an image merge field named "Logo".
            builder.InsertField("MERGEFIELD Image:Logo");

            // Set the callback that will supply the static logo.
            template.MailMerge.FieldMergingCallback = new StaticLogoCallback(logoPath);

            // Prepare a data source. The actual value is irrelevant because the callback ignores it.
            DataTable data = new DataTable("Data");
            data.Columns.Add("Logo");
            data.Rows.Add("ignored");

            // Perform the mail merge.
            template.MailMerge.Execute(data);

            // Save the merged document.
            string outputPath = Path.Combine(workDir, "MergedDocument.docx");
            template.Save(outputPath);
        }
    }
}
