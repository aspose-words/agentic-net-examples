using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertOleObjectToPdf
{
    static void Main()
    {
        // Paths to the source OLE file and the output PDF.
        string dataDir = @"C:\Data\";
        string oleFilePath = Path.Combine(dataDir, "Spreadsheet.xlsx"); // File to embed as OLE.
        string outputPdfPath = Path.Combine(dataDir, "DocumentWithOle.pdf");

        // Create a new empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the OLE object (embedded Excel spreadsheet) at the current cursor position.
        // Parameters: file name, ProgID, isLinked (false = embed), asIcon (false = show content), presentation (null = default icon).
        builder.InsertOleObject(oleFilePath, "Excel.Sheet", false, false, null);

        // Configure PDF save options to embed OLE objects as attachments.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as a PDF file with the specified options.
        doc.Save(outputPdfPath, pdfOptions);
    }
}
