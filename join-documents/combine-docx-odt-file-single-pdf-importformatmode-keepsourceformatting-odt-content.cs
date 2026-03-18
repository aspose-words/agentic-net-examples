using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the first document (DOCX) using the Document(string) constructor.
        Document docx = new Document("input.docx");

        // Load the second document (ODT) using the Document(string) constructor.
        Document odt = new Document("input.odt");

        // Append the ODT document to the DOCX document.
        // Use KeepSourceFormatting to preserve the original formatting of the ODT content.
        docx.AppendDocument(odt, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as a single PDF file.
        // The Save method with a file name and SaveFormat ensures the correct format is used.
        docx.Save("output.pdf", SaveFormat.Pdf);
    }
}
