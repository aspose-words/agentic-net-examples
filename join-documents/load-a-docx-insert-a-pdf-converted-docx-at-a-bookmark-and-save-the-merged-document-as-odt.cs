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

        // Paths for the sample documents.
        string mainDocPath = Path.Combine(outputDir, "MainDocument.docx");
        string pdfConvertedDocPath = Path.Combine(outputDir, "PdfConvertedDocument.docx");
        string mergedDocPath = Path.Combine(outputDir, "MergedDocument.odt");

        // -----------------------------------------------------------------
        // 1. Create the main DOCX document with a bookmark where we will insert.
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the original document.");
        mainBuilder.StartBookmark("InsertHere");
        mainBuilder.Writeln("Content before insertion (inside bookmark).");
        mainBuilder.EndBookmark("InsertHere");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a second DOCX that simulates a PDF‑to‑DOCX conversion.
        // -----------------------------------------------------------------
        Document pdfConvertedDoc = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfConvertedDoc);
        pdfBuilder.Writeln("This content originated from a PDF conversion.");
        pdfConvertedDoc.Save(pdfConvertedDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 3. Load the documents.
        // -----------------------------------------------------------------
        Document loadedMain = new Document(mainDocPath);
        Document loadedInsert = new Document(pdfConvertedDocPath);

        // -----------------------------------------------------------------
        // 4. Move to the bookmark and insert the second document.
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedMain);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(loadedInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Save the merged document as ODT.
        // -----------------------------------------------------------------
        loadedMain.Save(mergedDocPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // 6. Validation: ensure the file exists and contains text from both sources.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedDocPath))
            throw new InvalidOperationException("Merged ODT file was not created.");

        Document validationDoc = new Document(mergedDocPath);
        string mergedText = validationDoc.GetText();

        if (!mergedText.Contains("This is the original document.") ||
            !mergedText.Contains("This content originated from a PDF conversion."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Program completed successfully.
    }
}
