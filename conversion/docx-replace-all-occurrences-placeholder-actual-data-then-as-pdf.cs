using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Create a temporary folder for input and output files.
            string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeWordsExample");
            Directory.CreateDirectory(tempFolder);

            // Path to the temporary DOCX file.
            string inputPath = Path.Combine(tempFolder, "Template.docx");

            // Path to the output PDF file.
            string outputPath = Path.Combine(tempFolder, "Result.pdf");

            // Create a new document with placeholder text.
            Document doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Dear {{FirstName}} {{LastName}},");
            builder.Writeln("Your appointment is scheduled for {{Date}}.");
            builder.Writeln("Thank you!");

            // Save the template document to the temporary location.
            doc.Save(inputPath, SaveFormat.Docx);

            // Reload the document from the file (simulating a real-world scenario).
            Document loadedDoc = new Document(inputPath);

            // Replace placeholders with actual values.
            loadedDoc.Range.Replace("{{FirstName}}", "John");
            loadedDoc.Range.Replace("{{LastName}}", "Doe");
            loadedDoc.Range.Replace("{{Date}}", DateTime.Today.ToString("d"));

            // Save the modified document as PDF.
            loadedDoc.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF generated at: {outputPath}");
        }
    }
}
