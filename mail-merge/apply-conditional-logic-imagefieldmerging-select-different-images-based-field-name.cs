using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;                     // For Image.FromFile (NET Framework)
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsImageMergeExample
{
    // Callback that selects an image based on the merge field name.
    public class ConditionalImageCallback : IFieldMergingCallback
    {
        // Mapping of field names (without the "Image:" prefix) to image file paths.
        private readonly Dictionary<string, string> _fieldToImage = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Logo",      ImageDir + "Logo.jpg" },
            { "Signature", ImageDir + "Signature.png" },
            { "Stamp",     ImageDir + "Stamp.bmp" }
        };

        // Not used for plain text fields.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // No custom handling for non‑image fields.
        }

        // Called for each image merge field.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // args.DocumentFieldName contains the field name as written in the document,
            // e.g. "Image:Logo". We need the part after the prefix.
            string fieldName = args.FieldName; // Already stripped of the "Image:" prefix.

            if (_fieldToImage.TryGetValue(fieldName, out string imagePath))
            {
#if NET461_OR_GREATER || JAVA
                // Load the image from the file system.
                args.Image = Image.FromFile(imagePath);
#elif NET5_0_OR_GREATER
                // For .NET 5+ use SkiaSharp.
                args.Image = SkiaSharp.SKBitmap.Decode(imagePath);
                args.ImageFileName = imagePath;
#endif
            }
            else
            {
                // If no mapping is found, leave the image null – the merge will insert nothing.
                args.Image = null;
            }
        }

        // Helper to locate the images folder (adjust as needed).
        private static string ImageDir => @"C:\Images\"; // <-- change to your actual image directory.
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Insert two image merge fields: one for "Logo" and one for "Signature".
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Image:Logo");
            builder.Writeln(); // Add a line break between the images.
            builder.InsertField("MERGEFIELD Image:Signature");

            // Prepare a data source. The column names must match the field names without the "Image:" prefix.
            DataTable table = new DataTable("Images");
            table.Columns.Add("Logo", typeof(string));
            table.Columns.Add("Signature", typeof(string));

            // The actual values are irrelevant because the callback ignores them;
            // we just need rows to trigger the merge for each field.
            table.Rows.Add("unused1", "unused2");

            // Assign the custom callback.
            doc.MailMerge.FieldMergingCallback = new ConditionalImageCallback();

            // Execute the mail merge.
            doc.MailMerge.Execute(table);

            // Save the resulting document.
            doc.Save(@"C:\Output\ConditionalImageMerge.docx");
        }
    }
}
