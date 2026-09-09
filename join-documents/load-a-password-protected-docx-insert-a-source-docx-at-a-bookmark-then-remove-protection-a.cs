using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the documents used in the example.
        string protectedPath = Path.Combine(dataDir, "Protected.docx");
        string sourcePath = Path.Combine(dataDir, "Source.docx");
        string resultPath = Path.Combine(dataDir, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create a destination document, add a bookmark and protect it.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the protected document.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Bookmark placeholder – will be replaced.");
        destBuilder.EndBookmark("InsertHere");
        // Protect the document with a password.
        destDoc.Protect(ProtectionType.ReadOnly, "destPassword");
        destDoc.Save(protectedPath);

        // ---------------------------------------------------------------
        // 2. Create a source document that will be inserted at the bookmark.
        // ---------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document inserted at the bookmark.");
        srcDoc.Save(sourcePath);

        // ---------------------------------------------------------------
        // 3. Load the protected document using the correct password.
        // ---------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("destPassword");
        Document loadedDest = new Document(protectedPath, loadOptions);

        // ---------------------------------------------------------------
        // 4. Insert the source document at the bookmark location.
        // ---------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedDest);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------------------------------------------------------------
        // 5. Remove protection from the document.
        // ---------------------------------------------------------------
        loadedDest.Unprotect();

        // ---------------------------------------------------------------
        // 6. Save the final merged document.
        // ---------------------------------------------------------------
        loadedDest.Save(resultPath);

        // ---------------------------------------------------------------
        // 7. Simple validation – ensure the file exists and contains both texts.
        // ---------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        string resultText = loadedDest.GetText();

        if (!resultText.Contains("This is the protected document.") ||
            !resultText.Contains("This is the source document inserted at the bookmark."))
        {
            throw new InvalidOperationException("Result document does not contain expected content.");
        }

        // The example finishes without requiring any user interaction.
    }
}
