using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\example.docx";

        // Path where the converted document will be saved (PDF in this case).
        string outputPath = @"C:\Docs\example_converted.pdf";

        // Load the DOCX document using the Document(string) constructor.
        Document doc = new Document(inputPath);

        // Save the loaded document to the desired format.
        // The format is inferred from the file extension (.pdf).
        doc.Save(outputPath);
    }
}
