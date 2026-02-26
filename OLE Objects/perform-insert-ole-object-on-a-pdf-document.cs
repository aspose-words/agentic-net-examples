using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertOleObjectToPdf
{
    static void Main()
    {
        // Path to the folder that contains the source OLE file and where the PDF will be saved.
        string dataDir = @"C:\Data\";

        // The file to be embedded as an OLE object (e.g., an Excel spreadsheet).
        string oleFilePath = Path.Combine(dataDir, "Spreadsheet.xlsx");

        // Create a new empty Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a description paragraph.
        builder.Writeln("Embedded OLE object (Excel spreadsheet):");

        // Insert the OLE object as an icon. 
        // Parameters: file name, ProgID, isLinked (false = embedded), asIcon (true), presentation (null = default icon).
        builder.InsertOleObject(oleFilePath, "Excel.Sheet", false, true, null);

        // Prepare PDF save options to embed the OLE object as an attachment.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations
        };

        // Save the document as a PDF file with the specified options.
        string pdfOutputPath = Path.Combine(dataDir, "Result.pdf");
        doc.Save(pdfOutputPath, pdfOptions);
    }
}
