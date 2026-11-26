using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workDir);

        // Paths for the sample documents.
        string sourceDocxPath = Path.Combine(workDir, "SourceDocument.docx");
        string pdfConvertedDocxPath = Path.Combine(workDir, "PdfConvertedDocument.docx");
        string mergedOdtPath = Path.Combine(workDir, "MergedOutput.odt");

        // -----------------------------------------------------------------
        // Create the source DOCX with a bookmark where the PDF‑converted DOCX will be inserted.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the original source document.");
        srcBuilder.StartBookmark("InsertHere");
        srcBuilder.Writeln("[Insertion point]");
        srcBuilder.EndBookmark("InsertHere");
        sourceDoc.Save(sourceDocxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a DOCX that simulates a PDF that has been converted to DOCX.
        // -----------------------------------------------------------------
        Document pdfConvertedDoc = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfConvertedDoc);
        pdfBuilder.Writeln("This content originates from a PDF that was converted to DOCX.");
        pdfConvertedDoc.Save(pdfConvertedDocxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents.
        // -----------------------------------------------------------------
        Document mainDoc = new Document(sourceDocxPath);
        Document insertDoc = new Document(pdfConvertedDocxPath);

        // -----------------------------------------------------------------
        // Insert the PDF‑converted document at the bookmark.
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(mainDoc);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the merged document as ODT.
        // -----------------------------------------------------------------
        mainDoc.Save(mergedOdtPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // Validation: ensure the merged file exists and contains content from both sources.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedOdtPath))
        {
            throw new FileNotFoundException("Merged ODT file was not created.", mergedOdtPath);
        }

        Document mergedDoc = new Document(mergedOdtPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("This is the original source document.") ||
            !mergedText.Contains("This content originates from a PDF that was converted to DOCX."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content from both source files.");
        }

        // Cleanup (optional): delete temporary files.
        // Directory.Delete(workDir, true);
    }
}
