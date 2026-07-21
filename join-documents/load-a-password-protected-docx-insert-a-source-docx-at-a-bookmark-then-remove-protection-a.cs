using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading; // Needed for LoadOptions

public class Program
{
    public static void Main()
    {
        // Create an output folder for all temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string protectedPath = Path.Combine(outputDir, "protected.docx");
        string sourcePath = Path.Combine(outputDir, "source.docx");
        string mergedPath = Path.Combine(outputDir, "merged.docx");

        // -------------------------------------------------------------
        // 1. Create a destination document, add a bookmark, and protect it.
        // -------------------------------------------------------------
        Document protectedDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(protectedDoc);

        builder.Writeln("This is the beginning of the protected document.");
        builder.StartBookmark("InsertHere");               // Bookmark start
        builder.Writeln("Content that will be replaced."); // Placeholder
        builder.EndBookmark("InsertHere");                 // Bookmark end
        builder.Writeln("This is the end of the protected document.");

        // Apply read‑only protection with a password.
        protectedDoc.Protect(ProtectionType.ReadOnly, "SecretPwd");

        // Save the protected document.
        protectedDoc.Save(protectedPath);

        // -------------------------------------------------------------
        // 2. Create a source document that will be inserted at the bookmark.
        // -------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("=== Inserted Content Start ===");
        srcBuilder.Writeln("Hello from the source document!");
        srcBuilder.Writeln("=== Inserted Content End ===");
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------------------
        // 3. Load the protected document using the password.
        // -------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions("SecretPwd");
        Document loadedProtected = new Document(protectedPath, loadOptions);

        // -------------------------------------------------------------
        // 4. Insert the source document at the bookmark.
        // -------------------------------------------------------------
        DocumentBuilder insertBuilder = new DocumentBuilder(loadedProtected);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------------------
        // 5. Remove protection and save the merged result.
        // -------------------------------------------------------------
        loadedProtected.Unprotect(); // Removes protection regardless of password.
        loadedProtected.Save(mergedPath);

        // -------------------------------------------------------------
        // 6. Validate that the merged file was created.
        // -------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        Console.WriteLine($"Merged document saved to: {mergedPath}");
    }
}
