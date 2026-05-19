using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two image merge fields with different names.
        builder.InsertField("MERGEFIELD Image:Logo");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Image:Signature");

        // Prepare a dummy data source – the actual values are not used because the callback
        // decides which image to insert based on the field name.
        DataTable data = new DataTable("Images");
        data.Columns.Add("Logo");
        data.Columns.Add("Signature");
        data.Rows.Add("LogoPlaceholder", "SignaturePlaceholder");

        // Assign the custom callback that selects images conditionally.
        doc.MailMerge.FieldMergingCallback = new ConditionalImageCallback();

        // Perform the mail merge.
        doc.MailMerge.Execute(data);

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ConditionalImageMerge.docx");
        doc.Save(outputPath);
    }

    // Callback that chooses an image based on the merge field's full name.
    private class ConditionalImageCallback : IFieldMergingCallback
    {
        // No custom text handling needed.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args) { }

        // Called for each image merge field.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Determine which field is being merged.
            // args.DocumentFieldName contains the full field name, e.g., "Image:Logo".
            if (args.DocumentFieldName.Equals("Image:Logo", StringComparison.OrdinalIgnoreCase))
            {
                // Red square.
                args.ImageStream = GetRedSquare();
            }
            else if (args.DocumentFieldName.Equals("Image:Signature", StringComparison.OrdinalIgnoreCase))
            {
                // Blue rectangle.
                args.ImageStream = GetBlueRectangle();
            }
            else
            {
                // Gray square as fallback.
                args.ImageStream = GetGraySquare();
            }
        }

        // Returns a stream containing a 1x1 red PNG.
        private static MemoryStream GetRedSquare()
        {
            const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP8z8BQDwAF/AL+XKZcAAAAAElFTkSuQmCC";
            return CreateStreamFromBase64(base64);
        }

        // Returns a stream containing a 1x1 blue PNG.
        private static MemoryStream GetBlueRectangle()
        {
            const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP8z8DQAQAD/wJ/6VYVAAAAAElFTkSuQmCC";
            return CreateStreamFromBase64(base64);
        }

        // Returns a stream containing a 1x1 gray PNG.
        private static MemoryStream GetGraySquare()
        {
            const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP8z8DQAQAD/wJ/6VYVAAAAAElFTkSuQmCC";
            return CreateStreamFromBase64(base64);
        }

        // Helper to convert a base64 string to a MemoryStream positioned at the beginning.
        private static MemoryStream CreateStreamFromBase64(string base64)
        {
            byte[] bytes = Convert.FromBase64String(base64);
            var stream = new MemoryStream(bytes);
            stream.Position = 0;
            return stream;
        }
    }
}
