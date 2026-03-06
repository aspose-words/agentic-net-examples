using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOC file.
            string inputFile = @"C:\Input\SampleDocument.doc";

            // Path where the converted HTML file will be saved.
            string outputFile = @"C:\Output\SampleDocument.html";

            // Load the existing DOC document.
            // The Document constructor automatically detects the format based on the file extension.
            Document doc = new Document(inputFile);

            // Save the document in HTML format.
            // The Save method with a SaveFormat argument follows the provided rule.
            doc.Save(outputFile, SaveFormat.Html);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}
