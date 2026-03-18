using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

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
        public static void Main()
        {
            // Path to the Word template that contains an image merge field:
            //   <<Image:Photo>>
            // and a regular text merge field for the title:
            //   <<Title>>
            string templatePath = @"Templates\ReportTemplate.docx";

            // Prepare a collection of data sources, each representing a distinct report.
            var dataSources = new List<DataTable>();

            // -------------------------------------------------
            // Data source 1 – image supplied as a file path.
            // -------------------------------------------------
            var dt1 = new DataTable("Report");
            dt1.Columns.Add("Title", typeof(string));
            dt1.Columns.Add("Photo", typeof(string)); // file path
            dt1.Rows.Add("Quarterly Summary – Q1", @"Images\Logo1.jpg");
            dataSources.Add(dt1);

            // -------------------------------------------------
            // Data source 2 – image supplied as a byte array (e.g., from a database BLOB).
            // -------------------------------------------------
            var dt2 = new DataTable("Report");
            dt2.Columns.Add("Title", typeof(string));
            dt2.Columns.Add("Photo", typeof(byte[])); // raw bytes
            byte[] logoBytes = File.ReadAllBytes(@"Images\Logo2.png");
            dt2.Rows.Add("Quarterly Summary – Q2", logoBytes);
            dataSources.Add(dt2);

            // -------------------------------------------------
            // Data source 3 – another image file path.
            // -------------------------------------------------
            var dt3 = new DataTable("Report");
            dt3.Columns.Add("Title", typeof(string));
            dt3.Columns.Add("Photo", typeof(string)); // file path
            dt3.Rows.Add("Quarterly Summary – Q3", @"Images\Logo3.emf");
            dataSources.Add(dt3);

            // Instantiate the callback once – it will be reused for every document.
            var imageCallback = new ImageMergeCallback();

            // Process each data source and generate a separate report.
            for (int i = 0; i < dataSources.Count; i++)
            {
                // Load the template document (uses the provided load rule).
                Document doc = new Document(templatePath);

                // Attach the image handling callback.
                doc.MailMerge.FieldMergingCallback = imageCallback;

                // Execute mail merge. The overload that accepts a DataTable will merge all rows;
                // each DataTable contains only one row, producing a single report.
                doc.MailMerge.Execute(dataSources[i]);

                // Update any remaining fields (e.g., DATE fields) after the merge.
                doc.UpdateFields();

                // Save the populated report (uses the provided save rule).
                string outputPath = Path.Combine("Output", $"Report_{i + 1}.docx");
                doc.Save(outputPath);
            }
        }
    }
}
