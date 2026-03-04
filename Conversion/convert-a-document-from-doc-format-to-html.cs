using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOC file path
            string inputPath = @"C:\Input\Sample.doc";

            // Output HTML file path
            string outputPath = @"C:\Output\Sample.html";

            // Load the existing DOC document.
            // This uses the Document(string) constructor which automatically detects the format.
            Document doc = new Document(inputPath);

            // Save the document as HTML.
            // The Save method with (string, SaveFormat) follows the provided lifecycle rule.
            doc.Save(outputPath, SaveFormat.Html);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
