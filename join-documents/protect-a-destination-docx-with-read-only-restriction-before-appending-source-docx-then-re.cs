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
        string mergedPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedOutput.pdf");

        // -----------------------------------------------------------------
        // 1. Create the destination document and add some content.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");

        // Protect the destination document with a write‑protection password
        // and recommend it to be opened as read‑only.
        destDoc.WriteProtection.SetPassword("destPwd");
        destDoc.WriteProtection.ReadOnlyRecommended = true;

        // Save the protected destination document.
        destDoc.Save(destPath);

        // -----------------------------------------------------------------
        // 2. Create the source document and add some content.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the source document that will be appended.");

        // Save the source document.
        srcDoc.Save(srcPath);

        // -----------------------------------------------------------------
        // 3. Load both documents (demonstrating load from file) and append.
        // -----------------------------------------------------------------
        Document destination = new Document(destPath);
        Document source = new Document(srcPath);

        // Append the source document to the destination while keeping its formatting.
        destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 4. Remove the write‑protection before saving as PDF.
        // -----------------------------------------------------------------
        // Clear the password and the read‑only recommendation.
        destination.WriteProtection.SetPassword(string.Empty);
        destination.WriteProtection.ReadOnlyRecommended = false;

        // Save the merged document as PDF.
        destination.Save(mergedPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 5. Simple validation that the output files were created.
        // -----------------------------------------------------------------
        if (!File.Exists(destPath))
            throw new FileNotFoundException("Destination DOCX was not created.", destPath);
        if (!File.Exists(srcPath))
            throw new FileNotFoundException("Source DOCX was not created.", srcPath);
        if (!File.Exists(mergedPdfPath))
            throw new FileNotFoundException("Merged PDF was not created.", mergedPdfPath);

        // The program finishes without requiring any user interaction.
    }
}
