using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectInsertUnprotect
{
    static void Main()
    {
        const string password = "myPassword";
        const string resultPath = "Result.docx";

        // Create the destination document.
        Document destination = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destination);
        destBuilder.Writeln("This is the destination document.");

        // Protect the destination document with a password.
        destination.Protect(ProtectionType.ReadOnly, password);

        // Create the source document.
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);
        srcBuilder.Writeln("This is the source document.");

        // Insert the source document into the protected destination.
        DocumentBuilder builder = new DocumentBuilder(destination);
        builder.MoveToDocumentEnd();
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Remove protection from the combined document.
        destination.Unprotect(password);

        // Save the final document.
        destination.Save(resultPath, SaveFormat.Docx);
    }
}
