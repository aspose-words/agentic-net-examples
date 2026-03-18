using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class ProtectAndVerify
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DATE field – this will be the only editable element after protection.
        builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln(); // Move to a new line.

        // Insert regular text that should become read‑only for the user.
        builder.Writeln("This paragraph is not a field and should be protected.");

        // Apply document protection that allows only form fields (no form fields are present,
        // therefore the user cannot edit any regular content).
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        const string filePath = "ProtectedDocument.docx";
        doc.Save(filePath, SaveFormat.Docx);

        // Load the document back to verify protection.
        Document loadedDoc = new Document(filePath);

        // Verify that the document is indeed protected.
        if (loadedDoc.ProtectionType != ProtectionType.AllowOnlyFormFields)
            throw new InvalidOperationException("Document protection was not applied correctly.");

        // Attempt to modify non‑field content programmatically.
        // This is allowed by Aspose.Words, but the protection type ensures that a user
        // opening the document in Microsoft Word cannot edit this text.
        DocumentBuilder editBuilder = new DocumentBuilder(loadedDoc);
        editBuilder.MoveToDocumentEnd();
        editBuilder.Writeln("Programmatic edit: this line is added despite protection.");

        // Save the document after the programmatic edit.
        const string editedFilePath = "ProtectedDocument_Edited.docx";
        loadedDoc.Save(editedFilePath, SaveFormat.Docx);

        // Reload to confirm that protection remains intact after programmatic changes.
        Document finalDoc = new Document(editedFilePath);
        if (finalDoc.ProtectionType != ProtectionType.AllowOnlyFormFields)
            throw new InvalidOperationException("Document protection was lost after editing.");

        // Output verification results.
        Console.WriteLine("Document is protected with type: " + finalDoc.ProtectionType);
        Console.WriteLine("Non‑field content was edited programmatically, but remains protected for end users.");
    }
}
