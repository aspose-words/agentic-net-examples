using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Paths for the sample source documents.
        string docPath1 = Path.Combine(workDir, "Source1.docx");
        string docPath2 = Path.Combine(workDir, "Source2.docx");

        // Create the first sample DOCX.
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("First document content. This text is from source document 1.");
        srcDoc1.Save(docPath1, SaveFormat.Docx);

        // Create the second sample DOCX.
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Second document content. This text is from source document 2.");
        srcDoc2.Save(docPath2, SaveFormat.Docx);

        // Load the source documents.
        Document source1 = new Document(docPath1);
        Document source2 = new Document(docPath2);

        // Destination document that will hold the combined content.
        Document combined = new Document();

        // Append the first source document, preserving its original formatting.
        combined.AppendDocument(source1, ImportFormatMode.KeepSourceFormatting);

        // Append the second source document, also preserving its formatting.
        combined.AppendDocument(source2, ImportFormatMode.KeepSourceFormatting);

        // Path for the merged PDF output.
        string pdfPath = Path.Combine(workDir, "Combined.pdf");

        // Save the combined document as PDF.
        combined.Save(pdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The merged PDF file was not created.");

        // Load the PDF back into a Document to verify its text content.
        Document pdfDoc = new Document(pdfPath);
        string pdfText = pdfDoc.GetText();

        // Verify that both source texts are present in the PDF.
        if (!pdfText.Contains("First document content") || !pdfText.Contains("Second document content"))
            throw new InvalidOperationException("The merged PDF does not contain content from all source documents.");

        // Cleanup (optional): delete temporary files.
        // File.Delete(docPath1);
        // File.Delete(docPath2);
        // File.Delete(pdfPath);
    }
}
