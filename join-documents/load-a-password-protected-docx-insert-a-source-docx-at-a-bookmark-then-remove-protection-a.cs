using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Paths for sample files
        string sourcePath = "Source.docx";
        string protectedPath = "Protected.docx";
        string outputPath = "Result.docx";

        // 1. Create a source DOCX that will be inserted.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the content of the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // 2. Create a destination DOCX, add a bookmark, and protect it with a password.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the content of the destination document.");
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Bookmark placeholder.");
        destBuilder.EndBookmark("InsertHere");
        // Protect the document (read‑only) with password "SecretPwd".
        destDoc.Protect(ProtectionType.ReadOnly, "SecretPwd");
        destDoc.Save(protectedPath, SaveFormat.Docx);

        // 3. Load the password‑protected document using the password.
        LoadOptions loadOptions = new LoadOptions { Password = "SecretPwd" };
        Document protectedDoc = new Document(protectedPath, loadOptions);

        // 4. Load the source document to be inserted.
        Document insertDoc = new Document(sourcePath);

        // 5. Insert the source document at the bookmark location.
        DocumentBuilder insertBuilder = new DocumentBuilder(protectedDoc);
        insertBuilder.MoveToBookmark("InsertHere");
        insertBuilder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // 6. Remove protection from the merged document.
        protectedDoc.Unprotect();

        // 7. Save the final merged document.
        protectedDoc.Save(outputPath, SaveFormat.Docx);

        // 8. Validation: ensure the output file exists and contains source content.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The merged document was not saved.", outputPath);

        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();
        if (!resultText.Contains("This is the content of the source document."))
            throw new InvalidOperationException("The source content was not found in the merged document.");

        // Example completed successfully.
    }
}
