using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workDir);

        // Define file paths for the source documents and the merged PDF.
        string docxPath = Path.Combine(workDir, "SourceDocument.docx");
        string odtPath = Path.Combine(workDir, "SourceDocument.odt");
        string outputPdfPath = Path.Combine(workDir, "MergedOutput.pdf");

        // -----------------------------------------------------------------
        // Create a sample DOCX document.
        // -----------------------------------------------------------------
        var docxDocument = new Document();
        var docxBuilder = new DocumentBuilder(docxDocument);
        docxBuilder.Writeln("This is the DOCX document.");
        docxBuilder.Writeln("It contains some sample text.");
        docxDocument.Save(docxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample ODT document.
        // -----------------------------------------------------------------
        var odtDocument = new Document();
        var odtBuilder = new DocumentBuilder(odtDocument);
        odtBuilder.Writeln("This is the ODT document.");
        odtBuilder.Writeln("It will be appended with source formatting preserved.");
        odtDocument.Save(odtPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // Load the created documents.
        // -----------------------------------------------------------------
        var mainDoc = new Document(docxPath);
        var odtToAppend = new Document(odtPath);

        // Append the ODT document to the DOCX document, keeping source formatting.
        mainDoc.AppendDocument(odtToAppend, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the combined document as PDF.
        // -----------------------------------------------------------------
        mainDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validation: ensure the PDF file was created and contains content from both sources.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("The merged PDF file was not created.");
        }

        // Simple validation: the merged document should have at least two sections (one per source).
        if (mainDoc.Sections.Count < 2)
        {
            throw new InvalidOperationException("The merged document does not contain the expected number of sections.");
        }

        // Cleanup: optional removal of temporary files (comment out if inspection is needed).
        // File.Delete(docxPath);
        // File.Delete(odtPath);
        // File.Delete(outputPdfPath);
        // Directory.Delete(workDir, true);
    }
}
