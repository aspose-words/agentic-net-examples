using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary source and destination documents.
        const string destinationPath = "Destination.docx";
        const string sourcePath = "Source.docx";
        const string mergedPath = "Merged.docx";
        const string password = "pwd123";

        // ---------- Create destination document ----------
        Document destinationDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("This is the destination document.");
        // Protect the destination document with a password.
        destinationDoc.Protect(ProtectionType.ReadOnly, password);
        // Save the protected destination (optional, kept for completeness).
        destinationDoc.Save(destinationPath);

        // ---------- Create source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");
        sourceDoc.Save(sourcePath);

        // ---------- Load the protected destination document ----------
        // (Loading from file demonstrates that protection persists.)
        Document protectedDest = new Document(destinationPath);

        // Append the source document while preserving its formatting.
        protectedDest.AppendDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Remove protection using the correct password.
        protectedDest.Unprotect(password);

        // Save the merged result.
        protectedDest.Save(mergedPath);

        // Validate that the merged file was created.
        if (!File.Exists(mergedPath))
        {
            throw new InvalidOperationException($"Failed to create the merged document at '{mergedPath}'.");
        }

        // Optional: clean up temporary files (comment out if you want to inspect them).
        // File.Delete(destinationPath);
        // File.Delete(sourcePath);
    }
}
