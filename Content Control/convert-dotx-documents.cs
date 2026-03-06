using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOTX template from disk.
        // The Document(string) constructor automatically detects the format.
        Document template = new Document("Template.dotx");

        // Save the loaded template as a regular DOCX file.
        // Use the Save(string, SaveFormat) overload to specify the output format.
        template.Save("ConvertedDocument.docx", SaveFormat.Docx);

        // Additionally, save the same document as a PDF.
        // This demonstrates conversion to another common format.
        template.Save("ConvertedDocument.pdf", SaveFormat.Pdf);
    }
}
