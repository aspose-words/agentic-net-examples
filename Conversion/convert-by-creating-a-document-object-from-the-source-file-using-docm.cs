using System;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCM file.
        string sourceFile = @"C:\Docs\SourceDocument.docm";

        // Load the DOCM document. The constructor automatically detects the format.
        Document doc = new Document(sourceFile);

        // Verify that the document was loaded as a DOCM file.
        Console.WriteLine($"Original load format: {doc.OriginalLoadFormat}");

        // (Optional) Save the document in another format, e.g., DOCX.
        // doc.Save(@"C:\Docs\ConvertedDocument.docx", SaveFormat.Docx);
    }
}
