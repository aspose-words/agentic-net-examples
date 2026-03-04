using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace BatchConversionExample
{
    class Program
    {
        static void Main()
        {
            // Input Word document (any supported format, e.g., DOCX)
            string inputFile = @"C:\Input\SampleDocument.docx";

            // Output file path – the document will be saved in the legacy DOC format
            string outputFile = @"C:\Output\SampleDocumentConverted.doc";

            // Load the document using the Document(string) constructor (lifecycle rule)
            Document doc = new Document(inputFile);

            // Create DocSaveOptions specifying the target SaveFormat (Doc)
            DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);

            // Save the document using the Save(string, SaveOptions) overload (lifecycle rule)
            doc.Save(outputFile, saveOptions);

            Console.WriteLine("Document successfully converted to DOC format.");
        }
    }
}
