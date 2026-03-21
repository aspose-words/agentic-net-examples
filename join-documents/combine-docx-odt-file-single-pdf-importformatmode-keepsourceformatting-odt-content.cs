using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create the first document (DOCX) in memory.
        Document docx = new Document();
        DocumentBuilder builder = new DocumentBuilder(docx);
        builder.Writeln("This is the DOCX document.");

        // Create the second document (ODT) in memory.
        Document odt = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(odt);
        builder2.Writeln("This is the ODT document.");

        // Append the ODT document to the DOCX document, preserving source formatting.
        docx.AppendDocument(odt, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as a single PDF file.
        docx.Save("output.pdf", SaveFormat.Pdf);
    }
}
