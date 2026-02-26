using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string sourcePath = @"C:\Docs\Sample.doc";

        // Path to the target file (converted to PDF in this example).
        string targetPath = @"C:\Docs\Sample.pdf";

        // Load the existing DOC document. The constructor automatically detects the format.
        Document doc = new Document(sourcePath);

        // Save the document in the desired format. Here we convert to PDF.
        doc.Save(targetPath, SaveFormat.Pdf);
    }
}
