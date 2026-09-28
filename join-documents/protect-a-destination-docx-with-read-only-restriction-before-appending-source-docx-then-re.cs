using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workFolder);

        // Paths for the documents.
        string destPath = Path.Combine(workFolder, "Destination.docx");
        string srcPath = Path.Combine(workFolder, "Source.docx");
        string pdfPath = Path.Combine(workFolder, "Result.pdf");

        // -----------------------------------------------------------------
        // Create the destination DOCX and protect it with read‑only restriction.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the original content of the destination document.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // Apply read‑only protection with a password.
        const string protectionPassword = "myPassword";
        destDoc.Protect(ProtectionType.ReadOnly, protectionPassword);
        destDoc.Save(destPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the source DOCX that will be appended.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the content of the source document to be appended.");
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the protected destination document and the source document.
        // -----------------------------------------------------------------
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Append the source document to the destination.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Remove the read‑only protection.
        destination.Unprotect(protectionPassword);

        // Save the combined document as PDF.
        destination.Save(pdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The PDF output was not created.");
        }

        // Optional: clean up temporary files (comment out if inspection is needed).
        // File.Delete(destPath);
        // File.Delete(srcPath);
        // Directory.Delete(workFolder, true);
    }
}
