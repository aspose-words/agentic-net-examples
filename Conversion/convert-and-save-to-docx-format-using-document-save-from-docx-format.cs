using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load a source document. The source can be any format supported by Aspose.Words.
        // Replace the file path with the actual location of your input file.
        Document doc = new Document("input.pdf");

        // Save the document in DOCX format.
        // The SaveFormat.Docx enum value explicitly specifies the DOCX output format.
        doc.Save("output.docx", SaveFormat.Docx);
    }
}
