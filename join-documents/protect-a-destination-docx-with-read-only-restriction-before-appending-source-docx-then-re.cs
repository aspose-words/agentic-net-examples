using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the temporary documents.
        string destPath = Path.Combine(artifactsDir, "Destination.docx");
        string srcPath = Path.Combine(artifactsDir, "Source.docx");
        string pdfPath = Path.Combine(artifactsDir, "Merged.pdf");

        // -----------------------------------------------------------------
        // 1. Create the destination document and add some content.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");

        // Save it (optional, just to have a physical file).
        destDoc.Save(destPath);

        // -----------------------------------------------------------------
        // 2. Protect the destination document with a read‑only restriction.
        //    Use a password so we can later remove the protection.
        // -----------------------------------------------------------------
        const string protectionPassword = "pwd123";
        destDoc.Protect(ProtectionType.ReadOnly, protectionPassword);

        // -----------------------------------------------------------------
        // 3. Create the source document and add some content.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document that will be appended.");

        // Save the source document (optional).
        srcDoc.Save(srcPath);

        // -----------------------------------------------------------------
        // 4. Append the source document to the protected destination document.
        //    Keep the source formatting.
        // -----------------------------------------------------------------
        destDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Remove the read‑only protection.
        // -----------------------------------------------------------------
        destDoc.Unprotect(protectionPassword);

        // -----------------------------------------------------------------
        // 6. Save the combined document as PDF.
        // -----------------------------------------------------------------
        destDoc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 7. Simple validation: ensure the PDF file exists and contains both texts.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("PDF file was not created.");

        // Load the PDF back as a Document to verify its text.
        Document pdfDoc = new Document(pdfPath);
        string pdfText = pdfDoc.GetText();

        if (!pdfText.Contains("This is the destination document.") ||
            !pdfText.Contains("This is the source document that will be appended."))
        {
            throw new InvalidOperationException("Merged PDF does not contain expected content.");
        }

        // All done.
        Console.WriteLine("Documents merged and saved as PDF successfully.");
    }
}
