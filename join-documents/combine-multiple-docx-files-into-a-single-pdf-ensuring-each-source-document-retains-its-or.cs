using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for sample source documents.
        string inputFolder = "InputDocs";
        Directory.CreateDirectory(inputFolder);

        // Create first sample DOCX document.
        string doc1Path = Path.Combine(inputFolder, "Doc1.docx");
        var doc1 = new Document();
        var builder1 = new DocumentBuilder(doc1);
        builder1.Font.Name = "Arial";
        builder1.Font.Size = 14;
        builder1.Writeln("This is the first document. It uses Arial 14pt.");
        doc1.Save(doc1Path, SaveFormat.Docx);

        // Create second sample DOCX document.
        string doc2Path = Path.Combine(inputFolder, "Doc2.docx");
        var doc2 = new Document();
        var builder2 = new DocumentBuilder(doc2);
        builder2.Font.Name = "Times New Roman";
        builder2.Font.Size = 12;
        builder2.Writeln("This is the second document. It uses Times New Roman 12pt.");
        doc2.Save(doc2Path, SaveFormat.Docx);

        // Load the first document as the destination.
        var combinedDoc = new Document(doc1Path);

        // Load the second document to be appended.
        var docToAppend = new Document(doc2Path);

        // Append the second document while preserving its original formatting.
        combinedDoc.AppendDocument(docToAppend, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF.
        string outputPdfPath = "Combined.pdf";
        combinedDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validation: ensure the PDF file was created.
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException($"Failed to create the output PDF at '{outputPdfPath}'.");
        }

        // Validation: ensure the combined document contains sections from both source documents.
        int expectedSections = doc1.Sections.Count + doc2.Sections.Count;
        if (combinedDoc.Sections.Count != expectedSections)
        {
            throw new InvalidOperationException($"The combined document contains {combinedDoc.Sections.Count} sections, but {expectedSections} were expected.");
        }

        // Clean up sample documents (optional).
        // File.Delete(doc1Path);
        // File.Delete(doc2Path);
        // Directory.Delete(inputFolder, true);
    }
}
