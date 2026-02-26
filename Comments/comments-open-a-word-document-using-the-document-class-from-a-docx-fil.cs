using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be opened.
        const string inputPath = @"C:\Docs\SampleDocument.docx";

        // Load the existing Word document using the Document(string) constructor.
        Document doc = new Document(inputPath);

        // Example operation: output the document's text to the console.
        Console.WriteLine(doc.GetText());

        // (Optional) Save the document to verify that loading succeeded.
        // const string outputPath = @"C:\Docs\SampleDocument_Copy.docx";
        // doc.Save(outputPath);
    }
}
