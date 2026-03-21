using System;
using System.IO;
using Aspose.Words;

namespace AsposeWordsRevisionsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define output file path in the current directory.
            string outputFile = Path.Combine(Environment.CurrentDirectory, "CleanedDocument.docx");

            // Create a new document and add some tracked changes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start tracking revisions.
            doc.StartTrackRevisions("Author");

            // Add content that will be part of the revision.
            builder.Writeln("This is a sample paragraph with tracked changes.");

            // End tracking revisions.
            doc.StopTrackRevisions();

            // Accept all tracked changes (revisions) in the document.
            doc.AcceptAllRevisions();

            // Save the cleaned document to a new file.
            doc.Save(outputFile);

            Console.WriteLine($"Document saved successfully to: {outputFile}");
        }
    }
}
