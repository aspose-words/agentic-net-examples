using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string inputPath = @"C:\Docs\Template.docx";

            // Path to the output PDF file.
            string outputPath = @"C:\Docs\Result.pdf";

            // Load the existing DOCX document.
            Document doc = new Document(inputPath);

            // Replace placeholders with actual values.
            // Example placeholders: {{FirstName}}, {{LastName}}, {{Date}}
            doc.Range.Replace("{{FirstName}}", "John");
            doc.Range.Replace("{{LastName}}", "Doe");
            doc.Range.Replace("{{Date}}", DateTime.Today.ToString("d"));

            // Save the modified document as PDF.
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
