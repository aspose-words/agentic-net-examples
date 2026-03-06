using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = "source.docx";

        // Path to the output file (example: PDF).
        string outputPath = "converted.pdf";

        // Open the source document using the DOCX format.
        Document document = new Document(sourcePath);

        // Convert and save the document to the desired format.
        document.Save(outputPath, SaveFormat.Pdf);
    }
}
