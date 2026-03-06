using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (can be any format supported by Aspose.Words)
        string inputPath = @"C:\Input\source.pdf";

        // Path where the DOCX file will be saved
        string outputPath = @"C:\Output\result.docx";

        // Load the document from the input file
        Document doc = new Document(inputPath);

        // Save the loaded document in DOCX format
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
