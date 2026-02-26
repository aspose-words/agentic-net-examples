// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document from the file system.
        // The Document(string) constructor opens the file and detects its format.
        Document doc = new Document("MyDocument.docx");

        // Print the whole document to the default printer.
        // The Print() method uses the standard (no UI) print controller.
        doc.Print();
    }
}
