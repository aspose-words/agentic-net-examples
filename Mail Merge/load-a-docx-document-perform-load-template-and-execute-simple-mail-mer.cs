using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template from disk.
        Document doc = new Document("Template.docx");

        // Define the mail‑merge field names present in the template
        // and the corresponding values to insert.
        string[] fieldNames = { "FullName", "Address" };
        object[] fieldValues = { "John Doe", "123 Main St., Anytown" };

        // Execute a simple mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Prepare JPEG save options (default quality).
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);

        // Save the merged document as a JPEG image.
        doc.Save("MergedOutput.jpg", jpegOptions);
    }
}
