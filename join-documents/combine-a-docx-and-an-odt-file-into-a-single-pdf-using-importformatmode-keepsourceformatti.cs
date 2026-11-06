using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare working directory
        string workingDir = Path.Combine(Directory.GetCurrentDirectory(), "WorkingDir");
        Directory.CreateDirectory(workingDir);

        // Define file paths
        string docxPath = Path.Combine(workingDir, "sample.docx");
        string odtPath = Path.Combine(workingDir, "sample.odt");
        string pdfPath = Path.Combine(workingDir, "merged.pdf");

        // Create a sample DOCX document
        Document docx = new Document();
        DocumentBuilder builderDocx = new DocumentBuilder(docx);
        builderDocx.Writeln("This is content from the DOCX document.");
        docx.Save(docxPath, SaveFormat.Docx);

        // Create a sample ODT document
        Document odt = new Document();
        DocumentBuilder builderOdt = new DocumentBuilder(odt);
        builderOdt.Writeln("This is content from the ODT document.");
        odt.Save(odtPath, SaveFormat.Odt);

        // Load the created documents
        Document mainDoc = new Document(docxPath);
        Document odtDoc = new Document(odtPath);

        // Append ODT content to DOCX using KeepSourceFormatting
        mainDoc.AppendDocument(odtDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as PDF
        mainDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validation: ensure PDF was created and contains content from both sources
        if (!File.Exists(pdfPath))
            throw new Exception("Merged PDF file was not created.");

        FileInfo pdfInfo = new FileInfo(pdfPath);
        if (pdfInfo.Length == 0)
            throw new Exception("Merged PDF file is empty.");

        if (mainDoc.Sections.Count < 2)
            throw new Exception("Merged document does not contain both source sections.");

        // Optional: indicate success
        Console.WriteLine("Documents merged successfully into PDF at: " + pdfPath);
    }
}
