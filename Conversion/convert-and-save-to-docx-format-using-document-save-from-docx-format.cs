using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing document. The source can be any format supported by Aspose.Words.
        // Replace the path with the actual file you want to convert.
        Document doc = new Document("input.pdf");

        // Save the document in DOCX format.
        // The Save method overload with (string, SaveFormat) is used as defined in the Aspose.Words API.
        doc.Save("output.docx", SaveFormat.Docx);
    }
}
