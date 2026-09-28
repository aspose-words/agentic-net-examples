using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("This is the original paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Make some modifications to generate revisions.
        builder.Writeln("This paragraph is added while tracking.");
        builder.MoveToDocumentStart();
        builder.Write("Inserted at the beginning. ");

        // Delete a portion of text.
        builder.MoveToDocumentEnd();
        builder.Write("This text will be deleted.");
        // Note: DeleteDocument removes the whole document; kept for demonstration.
        // builder.DeleteDocument(); // Uncomment if you really want to delete the document.

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Accept all revisions.
        doc.AcceptAllRevisions();

        // Protect the document to prevent further editing without tracking.
        // ProtectionType enum is available directly in Aspose.Words namespace.
        doc.Protect(ProtectionType.ReadOnly, "password123");

        // Save the locked document.
        doc.Save("LockedDocument.docx");
    }
}
