using System;
using System.IO;
using Aspose.Words;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourcePath = Path.Combine("C:", "Data", "Source.docx");

            // Load the DOCX document using the Document constructor that accepts a file name.
            Document doc = new Document(sourcePath);

            // Example operation: retrieve the plain text of the loaded document.
            string documentText = doc.GetText();
            Console.WriteLine("Document text:");
            Console.WriteLine(documentText);

            // Optional: save the loaded document to another format (e.g., PDF).
            string outputPath = Path.Combine("C:", "Data", "Output.pdf");
            doc.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
