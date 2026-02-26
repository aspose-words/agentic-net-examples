using System;
using Aspose.Words;

class PdfToOtherFormats
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "input.pdf";

        // Load the PDF document.
        Document doc = new Document(pdfPath);

        // Save the document as a Word template (DOT) format.
        string dotPath = "output.dot";
        doc.Save(dotPath, SaveFormat.Dot);

        // Convert the same document to additional formats.

        // Save as DOCX (Office Open XML WordprocessingML).
        doc.Save("output.docx", SaveFormat.Docx);

        // Save as RTF (Rich Text Format).
        doc.Save("output.rtf", SaveFormat.Rtf);

        // Save as HTML.
        doc.Save("output.html", SaveFormat.Html);

        // Save as plain text.
        doc.Save("output.txt", SaveFormat.Text);
    }
}
