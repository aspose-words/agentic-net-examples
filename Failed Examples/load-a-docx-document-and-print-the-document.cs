// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOCX file to be loaded.
        string filePath = @"C:\Docs\Sample.docx";

        // Load the existing document from the specified file.
        Document doc = new Document(filePath);

        // Print the whole document using the default printer.
        doc.Print();
    }
}
