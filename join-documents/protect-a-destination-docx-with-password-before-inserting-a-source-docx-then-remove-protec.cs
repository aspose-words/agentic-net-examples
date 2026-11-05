using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current directory.
        string destPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docx");
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedResult.docx");
        string password = "Secret123";

        // -----------------------------------------------------------------
        // Create the destination document and add some content.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the destination document.");
        destDoc.Save(destPath, SaveFormat.Docx);

        // Protect the destination document with a password (read‑only protection).
        destDoc.Protect(ProtectionType.ReadOnly, password);

        // -----------------------------------------------------------------
        // Create the source document and add some content.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the source document.");
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Load the source document from file (simulating a typical workflow).
        Document loadedSource = new Document(sourcePath);

        // Append the source document to the protected destination document.
        destDoc.AppendDocument(loadedSource, ImportFormatMode.KeepSourceFormatting);

        // Remove protection from the merged document.
        destDoc.Unprotect(password);

        // Save the final merged document.
        destDoc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Merged document was not saved correctly.");
        }
    }
}
