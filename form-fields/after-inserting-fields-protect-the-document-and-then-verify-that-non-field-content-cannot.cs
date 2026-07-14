using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Path for the output documents.
        const string protectedPath = "ProtectedFormFields.docx";
        const string editedPath = "EditedAfterProtection.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert normal text before the form field.
        builder.Writeln("This is normal text before the form field.");

        // Insert a single text input form field.
        FormField textField = builder.InsertTextInput(
            "MyTextField",                     // field name
            TextFormFieldType.Regular,         // field type
            "",                                // format (none)
            "Enter value",                     // placeholder text
            0);                                // no length limit

        // Insert normal text after the form field.
        builder.Writeln("This is normal text after the form field.");

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // ---------- Verification ----------
        // 1. Verify that the document protection type is set correctly.
        if (doc.ProtectionType != ProtectionType.AllowOnlyFormFields)
            throw new InvalidOperationException("Document protection was not applied.");

        // 2. Verify that exactly one form field exists.
        FormFieldCollection fields = doc.Range.FormFields;
        if (fields.Count != 1)
            throw new InvalidOperationException($"Expected 1 form field, found {fields.Count}.");

        // 3. Verify that the existing field is the one we inserted.
        if (fields[0] != textField || fields[0].Name != "MyTextField")
            throw new InvalidOperationException("Form field verification failed.");

        // Save the protected document.
        doc.Save(protectedPath);
        Console.WriteLine($"Protected document saved to '{protectedPath}'.");

        // Reload the document to ensure the protection persists.
        Document loadedDoc = new Document(protectedPath);
        if (loadedDoc.ProtectionType != ProtectionType.AllowOnlyFormFields)
            throw new InvalidOperationException("Protection did not persist after reload.");

        // Attempt to edit non‑form‑field content programmatically.
        // (In the UI this would be blocked, but programmatic changes are allowed.)
        DocumentBuilder editBuilder = new DocumentBuilder(loadedDoc);
        editBuilder.MoveToDocumentStart();
        editBuilder.Write("Edited: ");

        // Save the edited version.
        loadedDoc.Save(editedPath);
        Console.WriteLine($"Document edited programmatically and saved to '{editedPath}'.");
        Console.WriteLine("Verification completed: document is protected for forms, and non‑field content remains separate from form fields.");
    }
}
