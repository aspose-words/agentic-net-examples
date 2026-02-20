using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace MailMergeToPdfExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template DOCX that contains MERGEFIELD tags.
            string templatePath = @"C:\Docs\Template.docx";

            // Load the DOCX document.
            Document doc = new Document(templatePath);

            // Define the merge field names that exist in the template.
            string[] fieldNames = { "FirstName", "LastName", "Message" };

            // Provide the corresponding values for a single record.
            object[] fieldValues = { "John", "Doe", "Hello! This message was created with Aspose.Words mail merge." };

            // Execute the mail merge. This populates the MERGEFIELD tags with the supplied data.
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Render DrawingML shapes directly (instead of using fallback shapes).
                DmlRenderingMode = DmlRenderingMode.DrawingML,

                // Optional: improve rendering quality.
                UseHighQualityRendering = true,
                UseAntiAliasing = true
            };

            // Path where the resulting PDF will be saved.
            string pdfPath = @"C:\Docs\Report.pdf";

            // Save the merged document as PDF using the configured options.
            doc.Save(pdfPath, pdfOptions);

            Console.WriteLine("Mail merge completed and PDF saved to: " + pdfPath);
        }
    }
}
