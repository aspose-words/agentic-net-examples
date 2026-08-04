using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;
using SkiaSharp; // Used for image creation on .NET 5+ platforms

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two image merge fields with distinct names.
        builder.InsertField("MERGEFIELD Image:Logo");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Image:Signature");

        // Prepare a minimal data source – the actual values are not used because the callback decides the image.
        DataTable data = new DataTable("Images");
        data.Columns.Add("Dummy");
        data.Rows.Add("x"); // single record

        // Assign the custom callback that selects images based on the field name.
        doc.MailMerge.FieldMergingCallback = new ConditionalImageCallback();

        // Perform the mail merge.
        doc.MailMerge.Execute(data);

        // Save the merged document.
        doc.Save("ConditionalImageMerge.docx");
    }

    // Callback that provides different images depending on the merge field name.
    private class ConditionalImageCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // No text merging required for this example.
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // The full field name (e.g., "Image:Logo") is available via DocumentFieldName.
            string fieldName = args.DocumentFieldName;

            SKBitmap bitmap;

            if (fieldName.Contains("Logo"))
            {
                // Create a blue rectangle for the "Logo" field.
                bitmap = CreateColoredBitmap(120, 60, SKColors.Blue);
            }
            else if (fieldName.Contains("Signature"))
            {
                // Create a green rectangle for the "Signature" field.
                bitmap = CreateColoredBitmap(150, 50, SKColors.Green);
            }
            else
            {
                // Fallback gray rectangle.
                bitmap = CreateColoredBitmap(80, 80, SKColors.Gray);
            }

            // Assign the generated image to the merge field.
            args.Image = bitmap;
        }

        // Helper method to create a solid‑color bitmap using SkiaSharp.
        private static SKBitmap CreateColoredBitmap(int width, int height, SKColor color)
        {
            SKBitmap bmp = new SKBitmap(width, height);
            using (SKCanvas canvas = new SKCanvas(bmp))
            {
                canvas.Clear(color);
            }
            return bmp;
        }
    }
}
