using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample DOCX files.
        string doc1Path = Path.Combine(outputDir, "Doc1.docx");
        string doc2Path = Path.Combine(outputDir, "Doc2.docx");

        // Create first sample document with distinct formatting.
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Font.Size = 24;
        builder1.Font.Color = Color.Blue;
        builder1.Writeln("First Document");
        builder1.Font.Size = 12;
        builder1.Font.Color = Color.Black;
        builder1.Writeln("This is the content of the first document.");
        doc1.Save(doc1Path, SaveFormat.Docx);

        // Create second sample document with different formatting.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Font.Size = 24;
        builder2.Font.Color = Color.Green;
        builder2.Writeln("Second Document");
        builder2.Font.Size = 12;
        builder2.Font.Color = Color.Black;
        builder2.Writeln("This is the content of the second document.");
        doc2.Save(doc2Path, SaveFormat.Docx);

        // Load the source documents.
        Document srcDoc1 = new Document(doc1Path);
        Document srcDoc2 = new Document(doc2Path);

        // Destination document that will hold the merged content.
        Document mergedDoc = new Document();

        // Append source documents while preserving their original formatting.
        mergedDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);
        mergedDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF.
        string pdfPath = Path.Combine(outputDir, "Merged.pdf");
        mergedDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file exists.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The merged PDF was not created.");

        // Load the PDF to verify it contains content from both source documents.
        Document pdfDoc = new Document(pdfPath);
        string pdfText = pdfDoc.GetText();

        if (pdfDoc.PageCount == 0)
            throw new InvalidOperationException("The merged PDF contains no pages.");

        if (!pdfText.Contains("First Document") || !pdfText.Contains("Second Document"))
            throw new InvalidOperationException("The merged PDF does not contain expected content from source documents.");

        // Indicate successful completion.
        Console.WriteLine("Documents merged and saved to PDF successfully.");
    }
}
