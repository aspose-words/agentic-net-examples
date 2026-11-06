using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary source documents and final output
        string sourceDocPath = "SourceDocument.docx";
        string pdfConvertedDocPath = "PdfConvertedDocument.docx";
        string outputPath = "MergedDocument.odt";

        // -----------------------------------------------------------------
        // Create the first DOCX with a bookmark where the second document will be inserted
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is the beginning of the source document.");
        builder.StartBookmark("InsertHere");
        builder.Writeln("Bookmark location – content will be inserted here.");
        builder.EndBookmark("InsertHere");
        builder.Writeln("This is the end of the source document.");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the second DOCX that simulates a PDF‑converted document
        // -----------------------------------------------------------------
        Document pdfConvertedDoc = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfConvertedDoc);
        pdfBuilder.Writeln("=== PDF Converted Content Start ===");
        pdfBuilder.Writeln("This text represents content that originated from a PDF file.");
        pdfBuilder.Writeln("=== PDF Converted Content End ===");
        pdfConvertedDoc.Save(pdfConvertedDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents from the saved files
        // -----------------------------------------------------------------
        Document mainDoc = new Document(sourceDocPath);
        Document insertDoc = new Document(pdfConvertedDocPath);

        // -----------------------------------------------------------------
        // Move the builder to the bookmark and insert the second document
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(mainDoc);

        // Verify that the bookmark exists
        if (mainDoc.Range.Bookmarks["InsertHere"] == null)
        {
            throw new Exception("Bookmark 'InsertHere' was not found in the source document.");
        }

        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the merged document as ODT
        // -----------------------------------------------------------------
        mainDoc.Save(outputPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // Validation: ensure the output file exists and contains expected content
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the merged ODT file at '{outputPath}'.");
        }

        string mergedText = mainDoc.GetText();
        if (!mergedText.Contains("This is the beginning of the source document.") ||
            !mergedText.Contains("This text represents content that originated from a PDF file.") ||
            !mergedText.Contains("This is the end of the source document."))
        {
            throw new Exception("Merged document does not contain expected content from both source files.");
        }

        // -----------------------------------------------------------------
        // Clean up temporary files (optional)
        // -----------------------------------------------------------------
        File.Delete(sourceDocPath);
        File.Delete(pdfConvertedDocPath);
    }
}
