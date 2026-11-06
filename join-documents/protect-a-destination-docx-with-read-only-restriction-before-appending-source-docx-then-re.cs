using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string destPath = "dest.docx";
        string sourcePath = "source.docx";
        string outputPdfPath = "output.pdf";

        // Create destination document with sample content
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // Protect the destination document with read‑only restriction
        destDoc.Protect(ProtectionType.ReadOnly, "pwd");

        // Create source document with sample content
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the source document (already in memory, but demonstrate loading)
        Document srcToAppend = new Document(sourcePath);

        // Append source document to the protected destination document
        destDoc.AppendDocument(srcToAppend, ImportFormatMode.KeepSourceFormatting);

        // Remove protection before saving as PDF
        destDoc.Unprotect();

        // Save the combined document as PDF
        destDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException("Failed to create the output PDF.");
        }

        // Clean up temporary files (optional)
        File.Delete(destPath);
        File.Delete(sourcePath);
    }
}
