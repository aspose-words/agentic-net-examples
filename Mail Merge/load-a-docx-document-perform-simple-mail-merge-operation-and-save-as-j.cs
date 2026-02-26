using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMailMergeToJpeg
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document doc = new Document("Template.docx");

            // Define the merge field names present in the template and the values to insert.
            string[] fieldNames = { "FullName", "Address", "City" };
            object[] fieldValues = { "James Bond", "MI5 Headquarters", "London" };

            // Perform a simple mail merge for a single record.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the merged document as a JPEG image.
            // ImageSaveOptions specifies the output format and can be used to control rendering options.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            doc.Save("MergedOutput.jpg", jpegOptions);
        }
    }
}
