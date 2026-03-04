using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // ------------------------------------------------------------
        // Save the document as a macro‑enabled DOCM file using OoxmlSaveOptions.
        // ------------------------------------------------------------
        // OoxmlSaveOptions allows us to specify the exact OOXML format.
        OoxmlSaveOptions docmOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        // Example of setting an additional option – you can customize more if needed.
        docmOptions.ExportGeneratorName = false; // do not embed Aspose.Words generator info
        // Save the document as DOCM.
        doc.Save("Output.docm", docmOptions);

        // ------------------------------------------------------------
        // Convert the same document to Excel (XLSX) and control worksheet naming.
        // ------------------------------------------------------------
        // XlsxSaveOptions controls how sections are mapped to worksheets.
        XlsxSaveOptions xlsxOptions = new XlsxSaveOptions();
        // Each section of the Word document will be saved as a separate worksheet.
        // This effectively controls the worksheet naming (one worksheet per section).
        xlsxOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;
        // Save the document as XLSX.
        doc.Save("Output.xlsx", xlsxOptions);
    }
}
