using System;
using Aspose.Words;

namespace AsposeWordsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOC file.
            string inputFile = @"C:\Docs\SourceDocument.doc";

            // Path where the HTML output will be saved.
            string outputFile = @"C:\Docs\ConvertedDocument.html";

            // Load the existing DOC file.
            Document doc = new Document(inputFile);

            // Save the document in HTML format.
            doc.Save(outputFile, SaveFormat.Html);
        }
    }
}
