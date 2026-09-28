using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "source.docx");
        string pdfConvertedPath = Path.Combine(Directory.GetCurrentDirectory(), "pdfConverted.docx");
        string mergedPath = Path.Combine(Directory.GetCurrentDirectory(), "merged.odt");

        // -----------------------------------------------------------------
        // Create a sample source DOCX with a bookmark where the PDF content will be inserted
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the main document.");
        sourceBuilder.StartBookmark("InsertHere");
        sourceBuilder.Writeln("Bookmark location (will be replaced).");
        sourceBuilder.EndBookmark("InsertHere");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create a sample DOCX that represents the PDF‑converted document
        // -----------------------------------------------------------------
        Document pdfConvertedDoc = new Document();
        DocumentBuilder pdfBuilder = new DocumentBuilder(pdfConvertedDoc);
        pdfBuilder.Writeln("Content from PDF converted document.");
        pdfConvertedDoc.Save(pdfConvertedPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents
        // -----------------------------------------------------------------
        Document mainDoc = new Document(sourcePath);
        Document insertDoc = new Document(pdfConvertedPath);

        // -----------------------------------------------------------------
        // Insert the PDF‑converted document at the bookmark
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(mainDoc);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Save the merged document as ODT
        // -----------------------------------------------------------------
        mainDoc.Save(mergedPath, SaveFormat.Odt);

        // -----------------------------------------------------------------
        // Validation: ensure the merged file exists and contains expected content
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
        {
            throw new FileNotFoundException("Merged ODT file was not created.", mergedPath);
        }

        // Load the merged document to verify content
        Document verificationDoc = new Document(mergedPath);
        string mergedText = verificationDoc.GetText();

        if (!mergedText.Contains("Content from PDF converted document."))
        {
            throw new InvalidOperationException("Merged document does not contain the inserted PDF content.");
        }

        // Program completed successfully
        Console.WriteLine("Documents merged and saved as ODT successfully.");
    }
}
