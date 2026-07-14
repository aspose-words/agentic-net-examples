using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading; // Needed for LoadOptions

public class Program
{
    public static void Main()
    {
        // File names used in the example.
        const string protectedPath = "Protected.docx";
        const string sourcePath = "Source.docx";
        const string resultPath = "Result.docx";
        const string password = "Secret123";

        // -----------------------------------------------------------------
        // 1. Create a source DOCX that will be inserted later.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Create a destination DOCX, add a bookmark, and protect it with a password.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Destination before bookmark.");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Destination after bookmark.");

        // Save the document with password protection.
        OoxmlSaveOptions protectOptions = new OoxmlSaveOptions
        {
            Password = password
        };
        destDoc.Save(protectedPath, protectOptions);

        // -----------------------------------------------------------------
        // 3. Load the password‑protected document using LoadOptions.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(password);
        Document protectedDoc = new Document(protectedPath, loadOptions);

        // Load the source document that will be inserted.
        Document docToInsert = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 4. Insert the source document at the bookmark.
        // -----------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(protectedDoc);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Remove protection from the document.
        // -----------------------------------------------------------------
        protectedDoc.Unprotect();

        // -----------------------------------------------------------------
        // 6. Save the merged result.
        // -----------------------------------------------------------------
        protectedDoc.Save(resultPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 7. Simple validation.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The merged document was not saved.");

        string resultText = protectedDoc.GetText();
        if (!resultText.Contains("This is the source document."))
            throw new InvalidOperationException("The source content was not inserted.");
    }
}
