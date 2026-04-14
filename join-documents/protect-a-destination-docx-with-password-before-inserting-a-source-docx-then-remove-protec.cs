using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary and final files
        string destPath = "Destination.docx";
        string srcPath = "Source.docx";
        string mergedPath = "Merged.docx";

        // -------------------------
        // Create destination document
        // -------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");

        // Save the destination so we have a physical file (optional but keeps the example clear)
        destDoc.Save(destPath, SaveFormat.Docx);

        // Protect the destination document with a password
        const string password = "destPassword";
        destDoc.Protect(ProtectionType.ReadOnly, password);

        // -------------------------
        // Create source document
        // -------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document.");

        // Save the source document (optional)
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // -------------------------
        // Append source to protected destination
        // -------------------------
        destDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------
        // Remove protection
        // -------------------------
        destDoc.Unprotect(); // No password needed for programmatic unprotect

        // -------------------------
        // Save merged document
        // -------------------------
        destDoc.Save(mergedPath, SaveFormat.Docx);

        // -------------------------
        // Validation
        // -------------------------
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not saved.");

        Document result = new Document(mergedPath);
        string resultText = result.GetText();

        if (!resultText.Contains("destination document") || !resultText.Contains("source document"))
            throw new InvalidOperationException("Merged document does not contain expected content.");

        // Clean up temporary files (optional)
        File.Delete(destPath);
        File.Delete(srcPath);
    }
}
