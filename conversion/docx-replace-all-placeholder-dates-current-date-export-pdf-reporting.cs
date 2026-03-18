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
            string outputPath = @"C:\Docs\Report.pdf";

            // Load the existing DOCX document.
            Document doc = new Document(inputPath);

            // Define the placeholder text that should be replaced with the current date.
            // Adjust this value to match the actual placeholder used in your template.
            const string placeholder = "{{Date}}";

            // Format the current date as a short date string (you can change the format as needed).
            string currentDate = DateTime.Now.ToString("d");

            // Replace all occurrences of the placeholder with the current date.
            // The Replace method works on the document's Range and replaces text case‑insensitively.
            doc.Range.Replace(placeholder, currentDate);

            // Save the modified document as a PDF.
            // Use the Save overload that accepts a file name and a SaveFormat enum value.
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
