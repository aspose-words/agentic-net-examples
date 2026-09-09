using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the sample files.
        string mainDocPath = Path.Combine(artifactsDir, "Main.docx");
        string pdfPath = Path.Combine(artifactsDir, "Sample.pdf");
        string convertedDocxPath = Path.Combine(artifactsDir, "ConvertedFromPdf.docx");
        string mergedOdtPath = Path.Combine(artifactsDir, "Merged.odt");

        // -----------------------------------------------------------------
        // 1. Create the main DOCX that contains a bookmark where we will insert.
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document.");
        mainBuilder.StartBookmark("InsertHere");
        mainBuilder.Writeln("Bookmark placeholder.");
        mainBuilder.EndBookmark("InsertHere");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a simple PDF file (using Aspose.Words) that we will later convert.
        // -----------------------------------------------------------------
        Document pdfSource = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfSource);
        pdfBuilder.Writeln("This is content from the PDF source.");
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 3. Load the PDF and save it as a DOCX (PDF‑to‑DOCX conversion).
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath); // Aspose.Words can load PDF.
        pdfDoc.Save(convertedDocxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 4. Load the main document again, move to the bookmark, and insert the converted DOCX.
        // -----------------------------------------------------------------
        Document mainDocLoaded = new Document(mainDocPath);
        DocumentBuilder insertBuilder = new DocumentBuilder(mainDocLoaded);
        insertBuilder.MoveToBookmark("InsertHere");

        Document docToInsert = new Document(convertedDocxPath);
        insertBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Save the merged document as ODT.
        // -----------------------------------------------------------------
        OdtSaveOptions odtOptions = new OdtSaveOptions(); // default options
        mainDocLoaded.Save(mergedOdtPath, odtOptions);

        // -----------------------------------------------------------------
        // 6. Validation: ensure the file exists and contains text from both sources.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedOdtPath))
            throw new Exception("Merged ODT file was not created.");

        Document mergedDoc = new Document(mergedOdtPath);
        string mergedText = mergedDoc.GetText();

        if (!mergedText.Contains("This is the main document.") ||
            !mergedText.Contains("This is content from the PDF source."))
        {
            throw new Exception("Merged document does not contain expected content.");
        }

        // Indicate successful completion.
        Console.WriteLine("Documents merged and saved as ODT successfully.");
    }
}
