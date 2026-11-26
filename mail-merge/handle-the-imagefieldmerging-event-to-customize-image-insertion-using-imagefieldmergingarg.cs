using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    // Callback class that handles both generic field merging and image field merging.
    private class MyFieldMergingCallback : IFieldMergingCallback
    {
        // A small PNG image encoded in base64.
        private static readonly byte[] ImageBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAUAAAAFCAYAAACNbyblAAAAHElEQVQI12P4" +
            "//8/w38GIAXDIBKE0DHxgljNBAAO9TXL0Y4OHwAAAABJRU5ErkJggg==");

        // Called for non‑image fields. No action needed for this example.
        public void FieldMerging(FieldMergingArgs args)
        {
            // No custom processing required for text fields in this scenario.
        }

        // Called specifically for image fields.
        public void ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // Ensure we are handling the intended merge field.
            if (args.FieldName == "MyImage")
            {
                // Provide the image data via a memory stream.
                args.ImageStream = new MemoryStream(ImageBytes);
                // Clear the file name to force insertion as an image.
                args.ImageFileName = null;
            }
        }
    }

    public static void Main()
    {
        // Create a new document and build the template with an image merge field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Below is the merged image:");
        builder.InsertField("MERGEFIELD MyImage \\* MERGEFORMAT");
        builder.Writeln();

        // Assign the field merging callback that handles image insertion.
        doc.MailMerge.FieldMergingCallback = new MyFieldMergingCallback();

        // Prepare a simple data source; the actual value is ignored by the callback.
        string[] fieldNames = { "MyImage" };
        object[] fieldValues = { "placeholder" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the output file was created and is not empty.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException("The merged document was not created correctly.");
        }
    }
}
