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
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.pdf");

        // -------------------------
        // Create the destination document.
        // -------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");

        // Apply read‑only protection with a password.
        const string password = "pwd123";
        destDoc.Protect(ProtectionType.ReadOnly, password);

        // Save the protected destination (optional, just to have a file on disk).
        destDoc.Save(destPath);

        // -------------------------
        // Create the source document.
        // -------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document that will be appended.");

        // Save the source document (optional).
        srcDoc.Save(srcPath);

        // -------------------------
        // Append the source document to the protected destination.
        // -------------------------
        destDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------
        // Remove the read‑only restriction.
        // -------------------------
        destDoc.Unprotect(password);

        // -------------------------
        // Save the final merged document as PDF.
        // -------------------------
        destDoc.Save(pdfPath, SaveFormat.Pdf);

        // Simple validation that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Clean up intermediate files (optional).
        File.Delete(destPath);
        File.Delete(srcPath);
    }
}
