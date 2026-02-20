using System;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the folder that contains the example DOCX file.
        string dataDir = @"C:\Docs\";

        // Full path to the DOCX document to be loaded.
        string docPath = Path.Combine(dataDir, "Example.docx");

        // Load the document. The Document constructor automatically detects the format (DOCX).
        Document doc = new Document(docPath);

        // Output the plain text content of the loaded document to the console.
        Console.WriteLine(doc.GetText());
    }
}
