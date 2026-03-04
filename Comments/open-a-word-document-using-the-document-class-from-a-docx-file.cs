using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be opened.
        string docPath = @"C:\Docs\SampleDocument.docx";

        // Load the existing Word document from the file system.
        // This uses the Document(string) constructor, which is the documented load rule.
        Document doc = new Document(docPath);

        // Example operation: output the document's text to the console.
        Console.WriteLine(doc.GetText());

        // (Optional) Save the document to a new file to verify that it was loaded correctly.
        // This uses the Document.Save(string) method, which follows the provided save rule.
        string outputPath = @"C:\Docs\SampleDocument_Copy.docx";
        doc.Save(outputPath);
    }
}
