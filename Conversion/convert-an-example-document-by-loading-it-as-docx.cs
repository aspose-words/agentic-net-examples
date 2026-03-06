using System;
using System.IO;
using Aspose.Words;

class LoadDocxExample
{
    static void Main()
    {
        // Path to the source DOCX file that will be loaded.
        string inputPath = @"C:\Docs\Example.docx";

        // Path where the loaded document will be saved after processing.
        string outputPath = @"C:\Docs\Example_Copy.docx";

        // Load the existing DOCX document using the Document(string) constructor.
        Document doc = new Document(inputPath);

        // (Optional) Perform any processing on the document here.
        // For demonstration we simply output the document's text to the console.
        Console.WriteLine("Document text:");
        Console.WriteLine(doc.GetText());

        // Save the document to a new file using the Save(string) method.
        doc.Save(outputPath);
    }
}
