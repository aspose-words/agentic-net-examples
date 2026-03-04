using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = @"C:\Input\Sample.doc";

        // Load the document. The constructor automatically detects the format (DOC in this case).
        Document doc = new Document(sourcePath);

        // Path to the output file. Here we convert the DOC to PDF.
        string outputPath = @"C:\Output\Sample.pdf";

        // Save the document in the desired format using the Save method overload that accepts a SaveFormat.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
