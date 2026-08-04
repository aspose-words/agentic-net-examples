using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;      // Needed for LoadOptions
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "JoinDocsExample");
        Directory.CreateDirectory(workDir);

        // File names
        string sourcePath = Path.Combine(workDir, "Source.docx");
        string protectedPath = Path.Combine(workDir, "Protected.docx");
        string resultPath = Path.Combine(workDir, "Result.docx");
        const string password = "SecretPwd";

        // -------------------------------------------------
        // 1. Create the source document that will be inserted.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the content of the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Create the destination document, add a bookmark,
        //    protect it and save it with a password.
        // -------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Start of the protected document.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("[Placeholder for insertion]");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("End of the protected document.");

        // Apply read‑only protection with a password
        destDoc.Protect(ProtectionType.ReadOnly, password);

        // Save with encryption so the file is password‑protected
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx) { Password = password };
        destDoc.Save(protectedPath, saveOptions);

        // -------------------------------------------------
        // 3. Load the password‑protected document.
        // -------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(password);
        Document loadedDoc = new Document(protectedPath, loadOptions);

        // -------------------------------------------------
        // 4. Insert the source document at the bookmark.
        // -------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedDoc);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // 5. Remove protection and save the final document.
        // -------------------------------------------------
        loadedDoc.Unprotect(); // Removes protection regardless of password
        loadedDoc.Save(resultPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 6. Validation: ensure the result file exists and contains both texts.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        string resultText = loadedDoc.GetText();

        if (!resultText.Contains("Start of the protected document.") ||
            !resultText.Contains("This is the content of the source document.") ||
            !resultText.Contains("End of the protected document."))
        {
            throw new InvalidOperationException("Result document does not contain expected content.");
        }

        // Optional cleanup
        // Directory.Delete(workDir, true);
    }
}
