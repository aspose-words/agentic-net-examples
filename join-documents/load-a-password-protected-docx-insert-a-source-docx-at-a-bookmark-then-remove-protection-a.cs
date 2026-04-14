using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "AsposeDemo");
        Directory.CreateDirectory(workDir);

        // Paths for the sample documents.
        string sourcePath = Path.Combine(workDir, "Source.docx");
        string protectedPath = Path.Combine(workDir, "Protected.docx");
        string resultPath = Path.Combine(workDir, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create the source document that will be inserted later.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create the destination document, add a bookmark, protect it,
        //    and save it with a password (encrypted DOCX).
        // -----------------------------------------------------------------
        Document protectedDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(protectedDoc);
        destBuilder.Writeln("Destination document - before bookmark.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("[Bookmark location]");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Destination document - after bookmark.");

        // Apply read‑only protection (optional).
        protectedDoc.Protect(ProtectionType.ReadOnly, "dummyPassword");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            Password = "SecretPwd" // Encrypt the file with this password.
        };
        protectedDoc.Save(protectedPath, saveOptions);

        // -----------------------------------------------------------------
        // 3. Load the password‑protected document using the correct password.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("SecretPwd");
        Document loadedDoc = new Document(protectedPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Insert the source document at the bookmark.
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedDoc);
        insertBuilder.MoveToBookmark("InsertHere");
        Document docToInsert = new Document(sourcePath);
        insertBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Remove protection (the document is now editable programmatically).
        // -----------------------------------------------------------------
        loadedDoc.Unprotect();

        // -----------------------------------------------------------------
        // 6. Save the merged result.
        // -----------------------------------------------------------------
        loadedDoc.Save(resultPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 7. Simple validation – ensure the file exists and contains expected text.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        Document verificationDoc = new Document(resultPath);
        string text = verificationDoc.GetText();

        if (!text.Contains("This is the content of the source document."))
            throw new InvalidOperationException("Source content was not inserted correctly.");
    }
}
