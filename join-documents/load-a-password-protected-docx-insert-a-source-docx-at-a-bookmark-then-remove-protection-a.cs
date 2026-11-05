using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workFolder);

        // Define file paths.
        string sourcePath = Path.Combine(workFolder, "source.docx");
        string destPath = Path.Combine(workFolder, "dest.docx");
        string outputPath = Path.Combine(workFolder, "merged.docx");

        // -------------------------------------------------
        // 1. Create a password‑protected source DOCX.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document content.");
        OoxmlSaveOptions sourceSaveOptions = new OoxmlSaveOptions { Password = "Secret123" };
        sourceDoc.Save(sourcePath, sourceSaveOptions);

        // -------------------------------------------------
        // 2. Create a destination DOCX that contains a bookmark.
        // -------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination document start.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("[Bookmark placeholder]");
        destBuilder.EndBookmark("InsertHere");
        destBuilder.Writeln("Destination document end.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 3. Load the documents (source requires the password).
        // -------------------------------------------------
        LoadOptions loadOpts = new LoadOptions { Password = "Secret123" };
        Document protectedSource = new Document(sourcePath, loadOpts);
        Document destination = new Document(destPath);

        // -------------------------------------------------
        // 4. Insert the source document at the bookmark.
        // -------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(destination);
        // MoveToBookmark returns true if the bookmark exists; we ignore the return value here.
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(protectedSource, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // 5. Remove any protection from the merged document.
        // -------------------------------------------------
        destination.Unprotect();

        // -------------------------------------------------
        // 6. Save the merged result.
        // -------------------------------------------------
        destination.Save(outputPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 7. Validate that the output file exists and contains the expected text.
        // -------------------------------------------------
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Merged document was not created.");

        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("This is the source document content."))
            throw new InvalidOperationException("Source content was not inserted correctly.");

        // Optional cleanup: uncomment the following line to delete temporary files.
        // Directory.Delete(workFolder, true);
    }
}
