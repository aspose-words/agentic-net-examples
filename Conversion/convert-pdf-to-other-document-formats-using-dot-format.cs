using System;
using Aspose.Words;

class PdfToOtherFormats
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "Input/sample.pdf";

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert the PDF to a Word template (DOT) format.
        string dotPath = "Output/sample.dot";
        pdfDoc.Save(dotPath, SaveFormat.Dot);

        // Load the generated DOT template.
        Document dotDoc = new Document(dotPath);

        // Convert the DOT template to various other formats.

        // DOCX (Office Open XML WordprocessingML Document)
        dotDoc.Save("Output/sample.docx", SaveFormat.Docx);

        // HTML (standard HTML format)
        dotDoc.Save("Output/sample.html", SaveFormat.Html);

        // RTF (Rich Text Format)
        dotDoc.Save("Output/sample.rtf", SaveFormat.Rtf);

        // ODT (OpenDocument Text)
        dotDoc.Save("Output/sample.odt", SaveFormat.Odt);
    }
}
