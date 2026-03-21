using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Replacing;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document in memory with a placeholder for the date.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This report was generated on {{Date}}.");

            // Define the placeholder text that should be replaced with the current date.
            const string placeholder = "{{Date}}";

            // Format the current date as a short date string (you can change the format as needed).
            string currentDate = DateTime.Now.ToString("d");

            // Replace all occurrences of the placeholder with the current date.
            var options = new FindReplaceOptions
            {
                Direction = FindReplaceDirection.Forward
            };
            doc.Range.Replace(placeholder, currentDate, options);

            // Determine an output path in the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.pdf");

            // Save the modified document as a PDF.
            doc.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF saved to: {outputPath}");
        }
    }
}
