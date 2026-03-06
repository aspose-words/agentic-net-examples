using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Template.docx");

        // Define the merge field names present in the template.
        string[] fieldNames = { "FullName", "Address", "City" };

        // Provide the corresponding values for a single record.
        object[] fieldValues = { "James Bond", "MI5 Headquarters", "London" };

        // Execute a simple mail merge for one record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Configure image save options to render the document as a JPEG.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);

        // Save the merged document as a JPEG image.
        doc.Save("MergedOutput.jpg", jpegOptions);
    }
}
