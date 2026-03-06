// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be loaded.
        string docPath = @"C:\Docs\Sample.docx";

        // Load the existing Word document from the file system.
        Document doc = new Document(docPath);

        // Print the entire document using the default printer.
        doc.Print();
    }
}
