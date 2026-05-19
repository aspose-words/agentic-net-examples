using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;      // Needed for LoadOptions
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the documents used in the example.
        string protectedDocPath = Path.Combine(dataDir, "Protected.docx");
        string sourceDocPath = Path.Combine(dataDir, "Source.docx");
        string resultDocPath = Path.Combine(dataDir, "Result.docx");

        // -----------------------------------------------------------------
        // 1. Create a password‑protected DOCX file.
        // -----------------------------------------------------------------
        Document protectedDoc = new Document();
        DocumentBuilder protectedBuilder = new DocumentBuilder(protectedDoc);
        protectedBuilder.Writeln("This is a password‑protected document.");

        // Insert a bookmark where the source document will be placed.
        protectedBuilder.StartBookmark("InsertHere");
        protectedBuilder.Writeln("[Bookmark content will be replaced]");
        protectedBuilder.EndBookmark("InsertHere");

        // Save the document with encryption password "Secret123".
        OoxmlSaveOptions protectSaveOptions = new OoxmlSaveOptions
        {
            Password = "Secret123"
        };
        protectedDoc.Save(protectedDocPath, protectSaveOptions);

        // -----------------------------------------------------------------
        // 2. Create a simple source DOCX that will be inserted.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This text comes from the source document.");
        sourceDoc.Save(sourceDocPath); // saved without password.

        // -----------------------------------------------------------------
        // 3. Load the protected document using the correct password.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("Secret123");
        Document mainDoc = new Document(protectedDocPath, loadOptions);

        // Load the source document that will be inserted.
        Document insertDoc = new Document(sourceDocPath);

        // -----------------------------------------------------------------
        // 4. Insert the source document at the bookmark.
        // -----------------------------------------------------------------
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);
        mainBuilder.MoveToBookmark("InsertHere");
        // Insert while keeping the source formatting.
        mainBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Remove protection from the document (if any).
        // -----------------------------------------------------------------
        mainDoc.Unprotect(); // Removes any protection regardless of password.

        // -----------------------------------------------------------------
        // 6. Save the merged result.
        // -----------------------------------------------------------------
        mainDoc.Save(resultDocPath);

        // -----------------------------------------------------------------
        // 7. Simple validation: ensure the file exists and contains expected text.
        // -----------------------------------------------------------------
        if (!File.Exists(resultDocPath))
            throw new InvalidOperationException("Result document was not created.");

        string resultText = mainDoc.GetText();
        if (!resultText.Contains("This text comes from the source document."))
            throw new InvalidOperationException("Source content was not inserted correctly.");

        // Example completed successfully.
    }
}
