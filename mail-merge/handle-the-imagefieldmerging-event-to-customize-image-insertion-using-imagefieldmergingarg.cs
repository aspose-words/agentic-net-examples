using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace AsposeWordsImageMergeExample
{
    // Implements the callback to handle image merge fields.
    public class ImageMergeHandler : IFieldMergingCallback
    {
        // No custom processing required for text fields.
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
        }

        // Called when the mail merge engine encounters an image merge field.
        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // The data source provides the image as a byte array.
            // Convert the byte array to a memory stream and assign it to the event args.
            if (args.FieldValue is byte[] imageBytes)
            {
                args.ImageStream = new MemoryStream(imageBytes);
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert an image merge field. The field name after the "Image:" prefix is "Photo".
            builder.InsertField("MERGEFIELD Image:Photo");

            // Assign the custom callback that will supply the image data.
            doc.MailMerge.FieldMergingCallback = new ImageMergeHandler();

            // Generate a simple in‑memory image (red pixel) and store it as a byte array.
            byte[] imageData = CreateSampleImage();

            // Perform the mail merge, supplying the image data for the "Photo" field.
            doc.MailMerge.Execute(
                new[] { "Photo" },               // Field names
                new object[] { imageData }       // Corresponding values
            );

            // Save the resulting document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedImage.docx");
            doc.Save(outputPath);
        }

        // Helper method that returns a small red PNG image as a byte array.
        private static byte[] CreateSampleImage()
        {
            // Base64‑encoded 1×1 red PNG image.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR4nGP8/5+hHgAHggJ/6VYVJwAAAABJRU5ErkJggg==";
            return Convert.FromBase64String(base64Png);
        }
    }
}
