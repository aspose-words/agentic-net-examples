using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string destPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string srcPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string mergedPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.pdf");

        // -------------------------------------------------
        // 1. Create the destination document and protect it.
        // -------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");

        // Apply write protection with a password and recommend read‑only.
        destDoc.WriteProtection.SetPassword("pwd123");
        destDoc.WriteProtection.ReadOnlyRecommended = true;

        // Save the protected destination document.
        destDoc.Save(destPath);

        // -------------------------------------------------
        // 2. Create the source document to be appended.
        // -------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document.");
        srcDoc.Save(srcPath);

        // -------------------------------------------------
        // 3. Load the protected destination document.
        // -------------------------------------------------
        Document loadedDest = new Document(destPath);

        // -------------------------------------------------
        // 4. Append the source document while keeping its formatting.
        // -------------------------------------------------
        loadedDest.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // 5. Remove write protection from the merged document.
        // -------------------------------------------------
        loadedDest.Unprotect();

        // -------------------------------------------------
        // 6. Save the final merged document as PDF.
        // -------------------------------------------------
        loadedDest.Save(mergedPdfPath, SaveFormat.Pdf);

        // -------------------------------------------------
        // 7. Simple validation that the files were created.
        // -------------------------------------------------
        if (!File.Exists(destPath))
            throw new FileNotFoundException("Destination DOCX was not created.", destPath);
        if (!File.Exists(srcPath))
            throw new FileNotFoundException("Source DOCX was not created.", srcPath);
        if (!File.Exists(mergedPdfPath))
            throw new FileNotFoundException("Merged PDF was not created.", mergedPdfPath);

        // Optional: verify that both texts are present in the merged document.
        string mergedText = loadedDest.GetText();
        if (!mergedText.Contains("This is the destination document.") ||
            !mergedText.Contains("This is the source document."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }
    }
}
