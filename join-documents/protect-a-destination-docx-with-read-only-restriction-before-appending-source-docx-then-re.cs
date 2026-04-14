using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary files.
        const string destinationPath = "Destination.docx";
        const string sourcePath = "Source.docx";
        const string mergedPdfPath = "Merged.pdf";

        // ---------- Create the destination document ----------
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("This is the destination document.");

        // Apply read‑only protection with a password.
        destinationDoc.Protect(ProtectionType.ReadOnly, "pwd");

        // Save the protected destination document.
        destinationDoc.Save(destinationPath);

        // ---------- Create the source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // ---------- Load the documents (simulating real files) ----------
        Document dest = new Document(destinationPath);
        Document src = new Document(sourcePath);

        // Append the source document to the destination, preserving source formatting.
        dest.AppendDocument(src, ImportFormatMode.KeepSourceFormatting);

        // Remove the read‑only protection after the append operation.
        dest.Unprotect("pwd"); // Password is optional; Unprotect() also works.

        // Save the merged document as PDF.
        dest.Save(mergedPdfPath, SaveFormat.Pdf);

        // ---------- Validation ----------
        if (!File.Exists(destinationPath) ||
            !File.Exists(sourcePath) ||
            !File.Exists(mergedPdfPath))
        {
            throw new InvalidOperationException("One or more output files were not created.");
        }
    }
}
