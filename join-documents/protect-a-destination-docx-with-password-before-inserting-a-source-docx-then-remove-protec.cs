using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string destPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string srcPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        const string password = "destPassword";

        // ---------- Create destination document ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");
        // Protect the destination document with a password.
        destDoc.Protect(ProtectionType.ReadOnly, password);
        // Save the protected destination (optional, just for illustration).
        destDoc.Save(destPath);

        // ---------- Create source document ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document to be inserted.");
        srcDoc.Save(srcPath);

        // ---------- Append source to protected destination ----------
        // Append while keeping source formatting.
        destDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------- Remove protection ----------
        // Unprotect using the correct password.
        bool unprotected = destDoc.Unprotect(password);
        if (!unprotected)
        {
            throw new InvalidOperationException("Failed to unprotect the document with the provided password.");
        }

        // ---------- Save the final merged document ----------
        destDoc.Save(resultPath, SaveFormat.Docx);

        // ---------- Validation ----------
        if (!File.Exists(resultPath))
        {
            throw new FileNotFoundException("The merged document was not created.", resultPath);
        }

        // Optional: verify that both pieces of text are present.
        Document finalDoc = new Document(resultPath);
        string text = finalDoc.GetText();
        if (!text.Contains("This is the destination document.") ||
            !text.Contains("This is the source document to be inserted."))
        {
            throw new InvalidOperationException("The merged document does not contain expected content.");
        }
    }
}
