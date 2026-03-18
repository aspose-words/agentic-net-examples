using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectInsertUnprotect
{
    static void Main()
    {
        // Paths to the documents.
        string destinationPath = "Destination.docx";
        string sourcePath = "Source.docx";
        string resultPath = "Result.docx";

        // Password to protect the destination document.
        const string password = "myPassword";

        // Load the destination document.
        Document destination = new Document(destinationPath);

        // Protect the destination document with a password (read‑only protection).
        destination.Protect(ProtectionType.ReadOnly, password);

        // Prepare a builder positioned at the end of the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);
        builder.MoveToDocumentEnd();

        // Load the source document to be inserted.
        Document source = new Document(sourcePath);

        // Insert the source document into the protected destination.
        // Keep the source formatting while inserting.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Remove protection from the combined document (password is not required for Unprotect()).
        destination.Unprotect();

        // Save the final document.
        destination.Save(resultPath, SaveFormat.Docx);
    }
}
