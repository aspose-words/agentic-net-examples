using System;
using Aspose.Words;

namespace AsposeWordsRevisionsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths.
            string inputFile = @"C:\Docs\OriginalWithRevisions.docx";
            string outputFile = @"C:\Docs\CleanedDocument.docx";

            // Load the existing DOCX document.
            Document doc = new Document(inputFile);

            // Accept all tracked changes (revisions) in the document.
            doc.AcceptAllRevisions();

            // Save the cleaned document to a new file.
            doc.Save(outputFile);
        }
    }
}
