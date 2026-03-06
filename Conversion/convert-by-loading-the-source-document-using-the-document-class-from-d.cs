using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Load the DOCX document using Aspose.Words Document constructor.
        Document doc = new Document(sourcePath);

        // Example: output the document's plain text to the console.
        Console.WriteLine(doc.GetText());

        // Optional: save the document to another format (e.g., PDF) using the save rule.
        // string outputPath = @"C:\Docs\Converted.pdf";
        // doc.Save(outputPath, SaveFormat.Pdf);
    }
}
