using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for sample documents and output
        string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");
        Directory.CreateDirectory(folderPath);

        // Paths for the source DOCX files and the merged PDF
        string docPath1 = Path.Combine(folderPath, "Doc1.docx");
        string docPath2 = Path.Combine(folderPath, "Doc2.docx");
        string outputPdfPath = Path.Combine(folderPath, "Combined.pdf");

        // Create first sample DOCX
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("This is the first document.");
        builder1.Font.Size = 16;
        builder1.Writeln("Header in first document.");
        doc1.Save(docPath1, SaveFormat.Docx);

        // Create second sample DOCX
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("This is the second document.");
        builder2.Font.Size = 14;
        builder2.Writeln("Different style in second document.");
        doc2.Save(docPath2, SaveFormat.Docx);

        // Load the source documents
        Document mainDoc = new Document(docPath1);
        Document appendDoc = new Document(docPath2);

        // Append the second document while keeping its original formatting
        mainDoc.AppendDocument(appendDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF
        mainDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The combined PDF file was not created.");
        }

        // Load the PDF to verify it contains content from both source documents
        Document resultDoc = new Document(outputPdfPath);
        if (resultDoc.PageCount < 2)
        {
            throw new InvalidOperationException("The combined PDF does not contain the expected content.");
        }
    }
}
