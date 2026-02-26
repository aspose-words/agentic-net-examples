using System;
using Aspose.Words;

namespace AsposeWordsDocToHtml
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOC file path – replace with your actual file location.
            string inputFile = @"C:\Docs\SampleDocument.doc";

            // Output HTML file path – the result will be saved here.
            string outputFile = @"C:\Docs\SampleDocument.html";

            // Load the existing DOC document.
            // The Document constructor automatically detects the format.
            Document doc = new Document(inputFile);

            // Save the document in HTML format using the Save method that accepts a SaveFormat enum.
            doc.Save(outputFile, SaveFormat.Html);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
