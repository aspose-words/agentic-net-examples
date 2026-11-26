using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary working directory
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeJoinDemo_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);

        // Paths for the documents
        string destDocPath = Path.Combine(workDir, "dest.docx");
        string sourceDocPath = Path.Combine(workDir, "source.docx");
        string mergedPdfPath = Path.Combine(workDir, "merged_encrypted.pdf");

        // 1. Create the destination DOCX (existing document)
        var destDoc = new Document();
        var destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the original document.");
        destDoc.Save(destDocPath, SaveFormat.Docx);

        // 2. Simulate obtaining a DOCX from a REST API by creating it locally
        var sourceDoc = new Document();
        var sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This content comes from the simulated REST API.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // 3. Load both documents
        var destination = new Document(destDocPath);
        var source = new Document(sourceDocPath);

        // 4. Append the source document to the destination document
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // 5. Save the merged document as an encrypted PDF
        var pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails(
                userPassword: "userPass123",
                ownerPassword: "ownerPass123")
        };
        destination.Save(mergedPdfPath, pdfOptions);

        // 6. Validation
        if (!File.Exists(mergedPdfPath))
        {
            throw new InvalidOperationException("The encrypted PDF was not created.");
        }

        // Verify that the merged document contains both sections
        if (destination.Sections.Count < 2)
        {
            throw new InvalidOperationException("The source document was not appended correctly.");
        }

        // Optional cleanup: uncomment to delete temporary files after execution
        // Directory.Delete(workDir, true);
    }
}
