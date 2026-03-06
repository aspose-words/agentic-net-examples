using System;
using Aspose.Words;

namespace DocToHtmlConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOC file path. Replace with the actual file location.
            string inputFile = @"C:\Input\Sample.doc";

            // Output HTML file path. The file will be created/overwritten.
            string outputFile = @"C:\Output\Sample.html";

            // Load the existing DOC document.
            Document doc = new Document(inputFile);

            // Save the document in HTML format.
            doc.Save(outputFile, SaveFormat.Html);
        }
    }
}
