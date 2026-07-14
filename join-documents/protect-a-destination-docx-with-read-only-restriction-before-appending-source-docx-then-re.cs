using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // File paths.
        string destPath = Path.Combine(outputDir, "Destination.docx");
        string srcPath = Path.Combine(outputDir, "Source.docx");
        string mergedPdfPath = Path.Combine(outputDir, "Merged.pdf");

        // ---------- Create destination document ----------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Destination document content.");

        // Apply read‑only write protection with a password.
        destDoc.WriteProtection.SetPassword("pwd");
        destDoc.WriteProtection.ReadOnlyRecommended = true;

        // Save the protected destination document (optional, just to have a file).
        destDoc.Save(destPath);

        // ---------- Create source document ----------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("Source document content.");
        srcDoc.Save(srcPath);

        // ---------- Append source to destination ----------
        destDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------- Remove write protection ----------
        destDoc.WriteProtection.SetPassword(string.Empty); // Clear password.
        destDoc.WriteProtection.ReadOnlyRecommended = false; // Clear read‑only flag.

        // ---------- Save merged document as PDF ----------
        destDoc.Save(mergedPdfPath, SaveFormat.Pdf);

        // ---------- Validation ----------
        if (!File.Exists(mergedPdfPath))
            throw new InvalidOperationException("Merged PDF was not created.");

        // (Optional) Clean up intermediate DOCX files if not needed.
        // File.Delete(destPath);
        // File.Delete(srcPath);
    }
}
