using System;
using Aspose.Words;

namespace DotToHtmlConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOT (Word template) file.
            string inputPath = @"C:\Input\Template.dot";

            // Path where the resulting HTML file will be saved.
            string outputPath = @"C:\Output\Template.html";

            // Load the DOT document. The constructor automatically detects the format.
            Document doc = new Document(inputPath);

            // Save the document as HTML using the explicit SaveFormat enumeration.
            doc.Save(outputPath, SaveFormat.Html);
        }
    }
}
