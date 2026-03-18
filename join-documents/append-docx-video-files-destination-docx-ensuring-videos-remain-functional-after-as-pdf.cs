using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Add a heading to the destination document.
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Combined Document with Videos");

        // Array of source DOCX file paths that already contain video objects
        // (online videos inserted via InsertOnlineVideo or OLE video objects).
        string[] sourceFiles = new string[]
        {
            @"C:\Docs\VideoDoc1.docx",
            @"C:\Docs\VideoDoc2.docx"
        };

        // Append each source document to the destination document.
        foreach (string srcPath in sourceFiles)
        {
            // Load the source document.
            Document srcDoc = new Document(srcPath);

            // Append while keeping the source formatting (including video shapes).
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Prepare PDF save options.
        // AttachmentsEmbeddingMode = Annotations embeds video files as PDF annotations,
        // which preserves their functionality when the PDF is opened.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AttachmentsEmbeddingMode = PdfAttachmentsEmbeddingMode.Annotations,
            // Update fields to ensure any video placeholders are up‑to‑date.
            UpdateFields = true
        };

        // Save the combined document as PDF using the configured options.
        dstDoc.Save(@"C:\Output\CombinedWithVideos.pdf", pdfOptions);
    }
}
