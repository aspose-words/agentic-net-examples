using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary and final documents.
        const string destinationPath = "Destination.docx";
        const string sourcePath = "Source.docx";
        const string mergedPath = "Merged.docx";

        // -----------------------------------------------------------------
        // 1. Create the destination document, add some content and protect it.
        // -----------------------------------------------------------------
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("This is the destination document.");
        // Protect the document with a password (read‑only protection).
        const string password = "pwd123";
        destinationDoc.Protect(ProtectionType.ReadOnly, password);
        // Save the protected document to disk.
        destinationDoc.Save(destinationPath);

        // -----------------------------------------------------------------
        // 2. Create the source document and add some content.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 3. Load the protected destination document and the source document.
        // -----------------------------------------------------------------
        // Protection does not encrypt the file, so it can be loaded without a password.
        Document dest = new Document(destinationPath);
        Document src = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 4. Append the source document to the destination document.
        // -----------------------------------------------------------------
        dest.AppendDocument(src, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Remove protection from the merged document.
        // -----------------------------------------------------------------
        // Unprotect works without providing the password, but the password can be supplied as well.
        dest.Unprotect();

        // -----------------------------------------------------------------
        // 6. Save the final merged document.
        // -----------------------------------------------------------------
        dest.Save(mergedPath);

        // -----------------------------------------------------------------
        // 7. Simple validation: ensure the merged file exists and contains both texts.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        string mergedText = dest.GetText();
        if (!mergedText.Contains("This is the destination document.") ||
            !mergedText.Contains("This is the source document."))
        {
            throw new InvalidOperationException("Merged document does not contain expected content.");
        }

        // Cleanup temporary files (optional).
        File.Delete(destinationPath);
        File.Delete(sourcePath);
    }
}
