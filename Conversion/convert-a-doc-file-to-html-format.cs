using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string inputFile = @"C:\Docs\Sample.doc";

        // Path where the HTML output will be saved.
        string outputFile = @"C:\Docs\Sample.html";

        // Load the DOC document from the file system.
        Document doc = new Document(inputFile);

        // Save the loaded document as HTML.
        doc.Save(outputFile, SaveFormat.Html);
    }
}
