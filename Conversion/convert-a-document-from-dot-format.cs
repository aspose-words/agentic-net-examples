using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDot
{
    static void Main()
    {
        // Load the DOT (Word template) file.
        Document doc = new Document("Template.dot");

        // Convert and save the document to DOCX format.
        doc.Save("Converted.docx", SaveFormat.Docx);
    }
}
