using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.BuildingBlocks;

namespace BatchReportGenerator
{
    // Callback that tells Aspose.Words how to handle image merge fields.
    // The field value can be either a file path (string) or a byte array.
    public class ImageMergeCallback : IFieldMergingCallback
    {
        public void FieldMerging(FieldMergingArgs args)
        {
            // No special handling for text fields in this scenario.
        }

        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // If the data source supplies a file path, use it directly.
            if (args.FieldValue is string path && File.Exists(path))
            {
                args.ImageFileName = path;
                return;
            }

            // If the data source supplies a byte array (e.g., from a DB BLOB), stream it.
            if (args.FieldValue is byte[] bytes)
            {
                args.ImageStream = new MemoryStream(bytes);
                return;
            }

            // If the value is null or unsupported, leave the field empty.
            args.ImageFileName = null;
            args.ImageStream = null;
        }
    }

    public class ReportGenerator
    {
        private static void EnsureDirectory(string path) => Directory.CreateDirectory(path);

        private static void CreatePlaceholderImage(string path)
        {
            // Minimal 1x1 pixel PNG (transparent)
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

        private static void PrepareResources()
        {
            // Ensure required folders exist
            EnsureDirectory("Images");
            EnsureDirectory("Templates");
            EnsureDirectory("Output");

            // Create placeholder image files if they don't exist
            string logo1Path = Path.Combine("Images", "Logo1.jpg");
            if (!File.Exists(logo1Path)) CreatePlaceholderImage(logo1Path);

            string logo2Path = Path.Combine("Images", "Logo2.png");
            if (!File.Exists(logo2Path)) CreatePlaceholderImage(logo2Path);

            string logo3Path = Path.Combine("Images", "Logo3.emf");
            if (!File.Exists(logo3Path)) CreatePlaceholderImage(logo3Path);

            // Create a simple template with merge fields if it doesn't exist
            string templatePath = Path.Combine("Templates", "ReportTemplate.docx");
            if (!File.Exists(templatePath))
            {
                Document templateDoc = new Document();
                var builder = new DocumentBuilder(templateDoc);

                builder.Writeln("Title:");
                builder.InsertField("MERGEFIELD Title \\* MERGEFORMAT");
                builder.Writeln();
                builder.Writeln("Photo:");
                // Insert an image merge field named Photo
                builder.InsertField("MERGEFIELD Photo \\* MERGEFORMAT");

                templateDoc.Save(templatePath);
            }
        }

        public static void Main()
        {
            PrepareResources();

            // Path to the Word template that contains an image merge field:
            //   <<Image:Photo>>
            // and a regular text merge field for the title:
            //   <<Title>>
            string templatePath = Path.Combine("Templates", "ReportTemplate.docx");

            // Prepare a collection of data sources, each representing a distinct report.
            var dataSources = new List<DataTable>();

            // -------------------------------------------------
            // Data source 1 – image supplied as a file path.
            // -------------------------------------------------
            var dt1 = new DataTable("Report");
            dt1.Columns.Add("Title", typeof(string));
            dt1.Columns.Add("Photo", typeof(string)); // file path
            dt1.Rows.Add("Quarterly Summary – Q1", Path.Combine("Images", "Logo1.jpg"));
            dataSources.Add(dt1);

            // -------------------------------------------------
            // Data source 2 – image supplied as a byte array (e.g., from a database BLOB).
            // -------------------------------------------------
            var dt2 = new DataTable("Report");
            dt2.Columns.Add("Title", typeof(string));
            dt2.Columns.Add("Photo", typeof(byte[])); // raw bytes
            byte[] logoBytes = File.ReadAllBytes(Path.Combine("Images", "Logo2.png"));
            dt2.Rows.Add("Quarterly Summary – Q2", logoBytes);
            dataSources.Add(dt2);

            // -------------------------------------------------
            // Data source 3 – another image file path.
            // -------------------------------------------------
            var dt3 = new DataTable("Report");
            dt3.Columns.Add("Title", typeof(string));
            dt3.Columns.Add("Photo", typeof(string)); // file path
            dt3.Rows.Add("Quarterly Summary – Q3", Path.Combine("Images", "Logo3.emf"));
            dataSources.Add(dt3);

            // Instantiate the callback once – it will be reused for every document.
            var imageCallback = new ImageMergeCallback();

            // Process each data source and generate a separate report.
            for (int i = 0; i < dataSources.Count; i++)
            {
                // Load the template document.
                Document doc = new Document(templatePath);

                // Attach the image handling callback.
                doc.MailMerge.FieldMergingCallback = imageCallback;

                // Execute mail merge. Each DataTable contains only one row, producing a single report.
                doc.MailMerge.Execute(dataSources[i]);

                // Update any remaining fields (e.g., DATE fields) after the merge.
                doc.UpdateFields();

                // Save the populated report.
                string outputPath = Path.Combine("Output", $"Report_{i + 1}.docx");
                doc.Save(outputPath);
            }

            Console.WriteLine("Reports generated successfully in the 'Output' folder.");
        }
    }
}
