using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeJoinDemo");
        Directory.CreateDirectory(workDir);

        // Paths for the sample source DOCX files.
        string docPath1 = Path.Combine(workDir, "Document1.docx");
        string docPath2 = Path.Combine(workDir, "Document2.docx");

        // Create first sample DOCX.
        var doc1 = new Document();
        var builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("First document - Heading");
        builder1.Font.Size = 12;
        builder1.Writeln("This is the content of the first document.");
        doc1.Save(docPath1, SaveFormat.Docx);

        // Create second sample DOCX.
        var doc2 = new Document();
        var builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Second document - Heading");
        builder2.Font.Size = 12;
        builder2.Writeln("This is the content of the second document.");
        doc2.Save(docPath2, SaveFormat.Docx);

        // Load the source documents.
        var sourceDoc1 = new Document(docPath1);
        var sourceDoc2 = new Document(docPath2);

        // Append the second document to the first, preserving source formatting.
        sourceDoc1.AppendDocument(sourceDoc2, ImportFormatMode.KeepSourceFormatting);

        // Path for the merged PDF output.
        string pdfPath = Path.Combine(workDir, "MergedOutput.pdf");

        // Save the combined document as PDF.
        sourceDoc1.Save(pdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("Merged PDF was not created.");
        }

        // Validation: ensure the merged document contains sections from both sources.
        if (sourceDoc1.Sections.Count < 2)
        {
            throw new InvalidOperationException("Merged document does not contain expected sections.");
        }

        // Cleanup: optional removal of temporary files (comment out if inspection is needed).
        // File.Delete(docPath1);
        // File.Delete(docPath2);
        // File.Delete(pdfPath);
        // Directory.Delete(workDir);
    }
}
