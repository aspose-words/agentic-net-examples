using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class VerifyProtectedFormFields
{
    static void Main()
    {
        // Prepare a folder for the test files.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string filePath = Path.Combine(artifactsDir, "ProtectedFormFields.docx");

        // ---------- Create and protect the document ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple form field.
        builder.Write("Please enter text here: ");
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Placeholder text", 0);

        // Protect the whole document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        doc.Save(filePath);

        // ---------- Load the document and verify protection ----------
        Document loadedDoc = new Document(filePath);

        // The document should still be protected for form fields.
        Debug.Assert(loadedDoc.ProtectionType == ProtectionType.AllowOnlyFormFields,
            "Document protection type is not AllowOnlyFormFields after reload.");

        // The section should be marked as protected for forms (default when document is protected).
        Debug.Assert(loadedDoc.Sections[0].ProtectedForForms,
            "Section is not marked as ProtectedForForms after reload.");

        // Attempt to edit the form field programmatically.
        FormField textInput = loadedDoc.Range.FormFields["TextInput1"];
        Debug.Assert(textInput != null, "Form field 'TextInput1' was not found.");

        // Set a new value for the form field.
        textInput.Result = "Edited text";

        // Verify that the new value was applied.
        Debug.Assert(textInput.Result == "Edited text",
            "Form field value was not updated correctly.");

        // Save the document again (optional, just to confirm no exceptions are thrown).
        loadedDoc.Save(filePath);
    }
}
