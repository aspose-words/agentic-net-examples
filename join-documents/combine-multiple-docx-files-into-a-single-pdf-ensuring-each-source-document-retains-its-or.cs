using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a working folder for temporary files.
        string workFolder = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocsWork");
        Directory.CreateDirectory(workFolder);

        // Paths for the sample source documents.
        string doc1Path = Path.Combine(workFolder, "Doc1.docx");
        string doc2Path = Path.Combine(workFolder, "Doc2.docx");

        // ---------- Create first sample DOCX ----------
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder1.Writeln("First Document Heading");
        builder1.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder1.Writeln("This is the first sample document.");
        doc1.Save(doc1Path, SaveFormat.Docx);

        // ---------- Create second sample DOCX ----------
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Font.Bold = true;
        builder2.Writeln("SECOND DOCUMENT TITLE");
        builder2.Font.Bold = false;
        builder2.Writeln("Content of the second document follows.");
        doc2.Save(doc2Path, SaveFormat.Docx);

        // Load the created documents.
        Document srcDoc1 = new Document(doc1Path);
        Document srcDoc2 = new Document(doc2Path);

        // Destination document – start with the first source document.
        // Use the file path constructor to avoid the Document‑to‑string conversion error.
        Document mergedDoc = new Document(doc1Path);

        // Append the second document while preserving its original formatting.
        mergedDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Define the output PDF path.
        string outputPdfPath = Path.Combine(workFolder, "Combined.pdf");

        // Save the merged document as PDF.
        mergedDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The merged PDF was not created.");

        // Optional cleanup (commented out for inspection).
        // File.Delete(doc1Path);
        // File.Delete(doc2Path);
        // File.Delete(outputPdfPath);
        // Directory.Delete(workFolder);
    }
}
