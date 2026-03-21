using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsImageMergeExample
{
    // Callback that provides an image for each merge field.
    public class ConditionalImageCallback : IFieldMergingCallback
    {
        // Not used for plain‑text fields.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args) { }

        // Called for each image merge field.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // A tiny placeholder PNG (16×16 transparent pixel) encoded in base64.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAF0lEQVR4AWP8z8Dwn4EIwDiqgAIAAP//AwA6VgZcAAAAAElFTkSuQmCC";

            // Write the PNG to a temporary file and tell Aspose.Words to use it.
            var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.png");
            File.WriteAllBytes(tempPath, Convert.FromBase64String(base64Png));

            args.ImageFileName = tempPath;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            var doc = new Document();

            // Insert two image merge fields: one for "Logo" and one for "Signature".
            var builder = new DocumentBuilder(doc);
            builder.InsertField("MERGEFIELD Image:Logo");
            builder.Writeln(); // Add a line break between the images.
            builder.InsertField("MERGEFIELD Image:Signature");

            // Prepare a data source. Column names must match the field names without the "Image:" prefix.
            var table = new DataTable("Images");
            table.Columns.Add("Logo", typeof(string));
            table.Columns.Add("Signature", typeof(string));
            table.Rows.Add("unused1", "unused2"); // Values are irrelevant; the callback supplies the images.

            // Assign the custom callback.
            doc.MailMerge.FieldMergingCallback = new ConditionalImageCallback();

            // Execute the mail merge.
            doc.MailMerge.Execute(table);

            // Save the resulting document in the current directory.
            doc.Save("ConditionalImageMerge.docx");
        }
    }
}
