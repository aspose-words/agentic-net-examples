using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "input.pdf";

        // Path where the resulting plain‑text file will be saved.
        string txtPath = "output.txt";

        // Load the PDF document.
        // ConvertShapeToOfficeMath = true ensures that any EquationXML shapes are turned into OfficeMath objects.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            ConvertShapeToOfficeMath = true
        };
        Document doc = new Document(pdfPath, loadOptions);

        // Save the document as plain text.
        // TxtSaveOptions.OfficeMathExportMode = Text (default) exports OfficeMath as Unicode characters.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            OfficeMathExportMode = TxtOfficeMathExportMode.Text
        };
        doc.Save(txtPath, saveOptions);
    }
}
