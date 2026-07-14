using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for temporary files.
        string sourceDocPath = Path.Combine(outputDir, "source.docx");
        string pdfPath = Path.Combine(outputDir, "temp.pdf");
        string pdfConvertedDocPath = Path.Combine(outputDir, "pdfConverted.docx");
        string mergedDocPath = Path.Combine(outputDir, "merged.odt");

        // ---------- Create the main DOCX with a bookmark ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the main document.");
        srcBuilder.StartBookmark("InsertHere");
        srcBuilder.Writeln("Bookmark start.");
        srcBuilder.EndBookmark("InsertHere");
        srcBuilder.Writeln("After bookmark.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // ---------- Create a PDF (simulating a PDF to be converted) ----------
        Document pdfSourceDoc = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfSourceDoc);
        pdfBuilder.Writeln("Content from PDF converted document.");
        pdfSourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // ---------- Load the PDF and save it as a DOCX (PDF‑converted DOCX) ----------
        Document pdfDoc = new Document(pdfPath); // Loads PDF and converts internally.
        pdfDoc.Save(pdfConvertedDocPath, SaveFormat.Docx);

        // ---------- Load the main document and insert the converted DOCX at the bookmark ----------
        Document mainDoc = new Document(sourceDocPath);
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.MoveToBookmark("InsertHere");

        Document insertDoc = new Document(pdfConvertedDocPath);
        mainBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------- Save the merged document as ODT ----------
        OdtSaveOptions odtOptions = new OdtSaveOptions();
        mainDoc.Save(mergedDocPath, odtOptions);

        // ---------- Validate that the merged ODT was created ----------
        if (!File.Exists(mergedDocPath))
            throw new InvalidOperationException("Merged ODT file was not created.");

        // Optional: clean up temporary files (comment out if inspection is needed).
        // File.Delete(sourceDocPath);
        // File.Delete(pdfPath);
        // File.Delete(pdfConvertedDocPath);
    }
}
