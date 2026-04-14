using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace ConditionalImageMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert two image merge fields with different logical names.
            builder.InsertField("MERGEFIELD Image:Logo");
            builder.Writeln(); // Add a line break between fields.
            builder.InsertField("MERGEFIELD Image:Signature");

            // Prepare a data source. The column names correspond to the logical names
            // used after the "Image:" prefix in the merge fields.
            DataTable data = new DataTable("Images");
            data.Columns.Add("Logo");
            data.Columns.Add("Signature");
            // The actual values are not used; they can be any placeholder.
            data.Rows.Add("Logo", "Signature");

            // Assign the custom callback that selects images based on the field name.
            doc.MailMerge.FieldMergingCallback = new ConditionalImageCallback();

            // Execute the mail merge.
            doc.MailMerge.Execute(data);

            // Save the resulting document.
            doc.Save("ConditionalImageMerge.docx");
        }

        // Implements IFieldMergingCallback to provide custom image selection logic.
        private class ConditionalImageCallback : IFieldMergingCallback
        {
            // Cache generated image streams so they are created only once.
            private readonly Dictionary<string, MemoryStream> _streamCache = new Dictionary<string, MemoryStream>();

            // This method is required by the interface but not used for image fields.
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                // No custom text merging needed.
            }

            // Called for each image merge field encountered during mail merge.
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // args.FieldName returns the name without the "Image:" prefix.
                string logicalName = args.FieldName;

                // Select or create the appropriate image stream based on the logical field name.
                if (logicalName == "Logo")
                {
                    args.ImageStream = GetOrCreateStream("Logo", RedPng);
                }
                else if (logicalName == "Signature")
                {
                    args.ImageStream = GetOrCreateStream("Signature", GreenPng);
                }
                else
                {
                    // Fallback: a simple blue placeholder.
                    args.ImageStream = GetOrCreateStream("Default", BluePng);
                }
            }

            // Retrieves a cached image stream or creates a new one from the supplied PNG bytes.
            private MemoryStream GetOrCreateStream(string key, byte[] pngBytes)
            {
                if (_streamCache.TryGetValue(key, out MemoryStream existing))
                {
                    // Reset position before reuse.
                    existing.Position = 0;
                    return existing;
                }

                MemoryStream stream = new MemoryStream(pngBytes);
                _streamCache[key] = stream;
                return stream;
            }

            // 1x1 pixel PNG images encoded as byte arrays.
            private static readonly byte[] RedPng = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z8BQDwAF/wJ/lKXUAAAAAElFTkSuQmCC");
            private static readonly byte[] GreenPng = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP4z8DAwAEAAf8C/6VbWQAAAABJRU5ErkJggg==");
            private static readonly byte[] BluePng = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP4z8DAwAEAAf8C/6VbWQAAAABJRU5ErkJggg==");
        }
    }
}
