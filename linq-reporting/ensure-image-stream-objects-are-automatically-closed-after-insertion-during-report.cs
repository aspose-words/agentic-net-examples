using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the template document that contains an image MERGEFIELD (e.g. MERGEFIELD Image:Photo).
        Document doc = new Document("Template.docx");

        // Prepare a data source. The Photo column holds raw image bytes.
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("FirstName", typeof(string));
        employees.Columns.Add("Photo", typeof(byte[]));

        // Add sample rows – replace the file paths with real images.
        employees.Rows.Add("John", File.ReadAllBytes("john.jpg"));
        employees.Rows.Add("Jane", File.ReadAllBytes("jane.png"));

        // Register the field‑merging callback. Aspose.Words will close the stream after the image is merged.
        doc.MailMerge.FieldMergingCallback = new ImageStreamCallback();

        // Execute the mail‑merge (or reporting) operation.
        doc.MailMerge.Execute(employees);

        // When saving to HTML we also want to guarantee that any streams created for images are closed.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ImageSavingCallback = new CloseImageStreamCallback();

        // Save the final document.
        doc.Save("Report.html", htmlOptions);
    }

    // Callback that provides an image stream for each image MERGEFIELD.
    // Aspose.Words closes the stream after the image is merged (ImageFieldMergingArgs.ImageStream).
    private class ImageStreamCallback : IFieldMergingCallback
    {
        void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
        {
            // No custom text handling required.
        }

        void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
        {
            // The data source supplies a byte[]; wrap it in a MemoryStream.
            // The stream will be disposed by Aspose.Words after the merge.
            if (args.FieldValue is byte[] imageBytes)
            {
                args.ImageStream = new MemoryStream(imageBytes);
            }
        }
    }

    // Callback that runs during HTML image saving.
    // Setting KeepImageStreamOpen to false forces Aspose.Words to close the stream after writing.
    private class CloseImageStreamCallback : IImageSavingCallback
    {
        void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
        {
            // Explicitly state that the stream should be closed after the image is saved.
            args.KeepImageStreamOpen = false;
        }
    }
}
