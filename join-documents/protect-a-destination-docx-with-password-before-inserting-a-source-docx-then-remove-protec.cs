using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string destPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string srcPath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        const string password = "destPwd";

        // -----------------------------------------------------------------
        // 1. Create the destination document, add some text, and protect it.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document. It will be protected before insertion.");
        // Protect the document with a password (read‑only protection).
        destDoc.Protect(ProtectionType.ReadOnly, password);
        // Save the protected destination (optional, just for demonstration).
        destDoc.Save(destPath);

        // -------------------------------------------------
        // 2. Create the source document with its own content.
        // -------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document. It will be inserted into the destination.");
        srcDoc.Save(srcPath);

        // --------------------------------------------------------------
        // 3. Load both documents (they are already in memory, but loading
        //    from the saved files demonstrates a realistic workflow).
        // --------------------------------------------------------------
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // --------------------------------------------------------------
        // 4. Append the source document to the protected destination.
        //    Use KeepSourceFormatting to preserve the source's appearance.
        // --------------------------------------------------------------
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // --------------------------------------------------------------
        // 5. Remove protection from the combined document.
        // --------------------------------------------------------------
        // Unprotect using the correct password.
        bool unprotected = destination.Unprotect(password);
        if (!unprotected)
        {
            throw new InvalidOperationException("Failed to unprotect the document with the provided password.");
        }

        // --------------------------------------------------------------
        // 6. Save the final merged document.
        // --------------------------------------------------------------
        destination.Save(resultPath);

        // --------------------------------------------------------------
        // 7. Validate that the result file was created.
        // --------------------------------------------------------------
        if (!File.Exists(resultPath))
        {
            throw new FileNotFoundException("The merged document was not saved correctly.", resultPath);
        }

        // Optional: output a simple confirmation (no interactive input required).
        Console.WriteLine("Document merged and saved to: " + resultPath);
    }
}
