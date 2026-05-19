using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // -----------------------------------------------------------------
        // 1. Create the main DOCX document with a bookmark where content will be inserted.
        // -----------------------------------------------------------------
        string mainDocPath = Path.Combine(dataDir, "Main.docx");
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.Writeln("This is the main document text.");
        mainBuilder.StartBookmark("InsertHere");
        mainBuilder.Writeln("[Placeholder for inserted content]");
        mainBuilder.EndBookmark("InsertHere");
        mainDoc.Save(mainDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a simple PDF document that will later be converted to DOCX.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(dataDir, "Source.pdf");
        Document pdfSource = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfSource);
        pdfBuilder.Writeln("This is the text extracted from the PDF.");
        pdfSource.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 3. Load the PDF and save it as a DOCX (PDF‑to‑DOCX conversion).
        // -----------------------------------------------------------------
        Document pdfAsDoc = new Document(pdfPath);
        string convertedDocxPath = Path.Combine(dataDir, "ConvertedFromPdf.docx");
        pdfAsDoc.Save(convertedDocxPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 4. Load the main document and the converted DOCX.
        // -----------------------------------------------------------------
        Document destination = new Document(mainDocPath);
        Document sourceToInsert = new Document(convertedDocxPath);

        // -----------------------------------------------------------------
        // 5. Insert the converted document at the bookmark.
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(destination);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(sourceToInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 6. Save the merged document as ODT.
        // -----------------------------------------------------------------
        string mergedOdtPath = Path.Combine(dataDir, "MergedDocument.odt");
        OdtSaveOptions odtOptions = new OdtSaveOptions();
        destination.Save(mergedOdtPath, odtOptions);

        // -----------------------------------------------------------------
        // 7. Validation: ensure the file exists and contains expected content.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedOdtPath))
            throw new InvalidOperationException("The merged ODT file was not created.");

        string mergedText = destination.GetText();
        if (!mergedText.Contains("This is the main document text.") ||
            !mergedText.Contains("This is the text extracted from the PDF."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Optional: indicate successful completion.
        Console.WriteLine("Document merged and saved as ODT successfully.");
    }
}
