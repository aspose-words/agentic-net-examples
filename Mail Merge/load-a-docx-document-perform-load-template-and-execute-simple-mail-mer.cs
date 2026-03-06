using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMailMergeToJpeg
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX template that contains MERGEFIELDs.
            const string templatePath = "Template.docx";

            // Path where the rendered JPEG image will be saved.
            const string outputPath = "Result.jpg";

            // Load the existing DOCX document.
            Document doc = new Document(templatePath);

            // Define the mail‑merge field names present in the template and the values to insert.
            string[] fieldNames = { "FullName", "Company" };
            object[] fieldValues = { "John Doe", "Acme Corp" };

            // Execute a simple mail merge for a single record.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Configure image save options to render the document as a JPEG.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);

            // Save the merged document as a JPEG image.
            doc.Save(outputPath, jpegOptions);
        }
    }
}
