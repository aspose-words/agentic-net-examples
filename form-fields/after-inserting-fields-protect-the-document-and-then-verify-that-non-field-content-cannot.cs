using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph that will become read‑only after protection.
        builder.Writeln("This paragraph should become read‑only after protection.");

        // Insert a text input form field. The last parameter is maxLength (0 = unlimited).
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default text", 0);

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBox", false, 0);

        // Insert a dropdown (combo box) form field with two items.
        string[] items = { "Option 1", "Option 2" };
        builder.InsertComboBox("DropDown", items, 0);

        // Save the unprotected version (optional).
        doc.Save("Unprotected.docx");

        // Protect the document for read‑only editing (blocks non‑field changes).
        // Older Aspose.Words versions may not expose ProtectionType.Forms, so ReadOnly is used.
        doc.Protect(ProtectionType.ReadOnly, "myPassword");

        // Save the protected document.
        doc.Save("Protected.docx");

        // Attempt to edit non‑field content after protection.
        bool editSucceeded = false;
        try
        {
            // Move the builder to the end of the document and try to add a new paragraph.
            builder.MoveToDocumentEnd();
            builder.Writeln("Attempting to edit after protection.");
            editSucceeded = true; // If no exception, edit succeeded (unexpected).
        }
        catch (Exception ex)
        {
            // Expected path: editing should be blocked.
            Console.WriteLine("Editing blocked as expected: " + ex.Message);
        }

        if (editSucceeded)
        {
            Console.WriteLine("Unexpectedly succeeded in editing protected document.");
        }

        // Verify that at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were found in the document.");
        }

        // Read and display the default value of the text input field.
        FormField textField = doc.Range.FormFields["TextField"];
        if (textField == null)
            throw new InvalidOperationException("TextField not found.");

        Console.WriteLine("TextField default value: " + textField.Result);
    }
}
