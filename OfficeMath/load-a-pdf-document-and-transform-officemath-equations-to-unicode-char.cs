using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfOfficeMathToUnicode
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\document.pdf";

        // Load the PDF with conversion of EquationXML shapes to OfficeMath objects.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true; // Enable conversion to OfficeMath.
        Document doc = new Document(pdfPath, loadOptions);

        // Prepare save options to export OfficeMath as plain Unicode text.
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.OfficeMathExportMode = TxtOfficeMathExportMode.Text; // Export as plain text.

        // Save the document to a memory stream using the TXT options.
        using (MemoryStream txtStream = new MemoryStream())
        {
            doc.Save(txtStream, saveOptions);
            txtStream.Position = 0;

            // Read the resulting Unicode text.
            string unicodeText = new StreamReader(txtStream, Encoding.UTF8).ReadToEnd();

            // Example usage: write the text to a file or process it further.
            File.WriteAllText(@"C:\Output\document.txt", unicodeText);
            Console.WriteLine("OfficeMath equations have been converted to Unicode text.");
        }
    }
}
