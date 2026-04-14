using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string workDir = Directory.GetCurrentDirectory();
        string sourceDocPath = Path.Combine(workDir, "Source.docx");
        string pdfPath = Path.Combine(workDir, "Source.pdf");
        string convertedDocPath = Path.Combine(workDir, "ConvertedFromPdf.docx");
        string mergedOutputPath = Path.Combine(workDir, "MergedDocument.odt");

        // -----------------------------------------------------------------
        // 1. Create the main DOCX document with a bookmark where content will be inserted.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the main document.");
        srcBuilder.StartBookmark("InsertHere");
        srcBuilder.Writeln("[Placeholder for inserted content]");
        srcBuilder.EndBookmark("InsertHere");
        sourceDoc.Save(sourceDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a PDF file that will later be converted to DOCX.
        // -----------------------------------------------------------------
        Document pdfSourceDoc = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfSourceDoc);
        pdfBuilder.Writeln("This is the content that originally came from a PDF.");
        pdfSourceDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 3. Load the PDF and convert it to DOCX.
        // -----------------------------------------------------------------
        Document pdfLoaded = new Document(pdfPath);
        pdfLoaded.Save(convertedDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 4. Load the main document and the converted DOCX, then insert at the bookmark.
        // -----------------------------------------------------------------
        Document mainDoc = new Document(sourceDocPath);
        Document insertDoc = new Document(convertedDocPath);

        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        // Move the cursor to the bookmark where the content should be inserted.
        mainBuilder.MoveToBookmark("InsertHere");
        // Insert the converted document inline, preserving its original formatting.
        mainBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Save the merged document as ODT.
        // -----------------------------------------------------------------
        // Using default ODT save options; could also instantiate OdtSaveOptions if needed.
        mainDoc.Save(mergedOutputPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // 6. Validation: ensure the output file exists and contains expected text.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedOutputPath))
            throw new InvalidOperationException("Merged ODT file was not created.");

        // Load the saved ODT to verify its content.
        Document resultDoc = new Document(mergedOutputPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("This is the main document.") ||
            !resultText.Contains("This is the content that originally came from a PDF."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // All done. No interactive prompts; the program will exit automatically.
    }
}
