using System;
using Aspose.Words;

namespace AsposeWordsConversionSample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input file path (any format supported by Aspose.Words, e.g., DOCX)
            string inputFile = @"C:\Docs\SampleDocument.docx";

            // Output file path – we want to convert to the legacy DOC format
            string outputFile = @"C:\Docs\SampleDocumentConverted.doc";

            // Load the source document using the Document constructor that accepts a file name
            Document doc = new Document(inputFile);

            // Save the document in the DOC format using the Save method that takes a file name and a SaveFormat enum
            doc.Save(outputFile, SaveFormat.Doc);

            Console.WriteLine("Document has been successfully converted to DOC format.");
        }
    }
}
