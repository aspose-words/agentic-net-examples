using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a normal paragraph (non‑form field content).
        builder.Writeln("This is a normal paragraph that should be protected.");

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput(
            "NameField",
            TextFormFieldType.Regular,
            "",
            "John Doe",
            0);
        builder.Writeln(); // Move to next line.

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptCheckBox", false, 0);
        builder.Writeln();

        // Insert a dropdown (combo box) form field.
        builder.Write("Select an option: ");
        string[] items = { "Option 1", "Option 2", "Option 3" };
        FormField comboBox = builder.InsertComboBox("OptionComboBox", items, 0);
        builder.Writeln();

        // Save the document before protection (optional).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFieldsProtected.docx");
        doc.Save(outputPath);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Verify that the document is protected for forms.
        if (doc.ProtectionType != ProtectionType.AllowOnlyFormFields)
            throw new InvalidOperationException("Document protection was not applied correctly.");

        // Verify that non‑field content is not a form field.
        // The first paragraph should not appear in the FormFields collection.
        bool nonFieldIsFormField = false;
        foreach (FormField ff in doc.Range.FormFields)
        {
            if (ff.Result == "This is a normal paragraph that should be protected.")
            {
                nonFieldIsFormField = true;
                break;
            }
        }

        if (nonFieldIsFormField)
            throw new InvalidOperationException("Non‑form field content was incorrectly identified as a form field.");

        // Attempt to modify the non‑form field text programmatically.
        // This change is allowed programmatically, but the document remains protected.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs.Clear();
        firstParagraph.AppendChild(new Run(doc, "Attempted modification of protected text."));

        // Save the protected document.
        doc.Save(outputPath);

        // Load the saved document to confirm protection persists.
        Document loadedDoc = new Document(outputPath);
        if (loadedDoc.ProtectionType != ProtectionType.AllowOnlyFormFields)
            throw new InvalidOperationException("Protection type was not persisted after saving.");

        // Output a simple confirmation to the console.
        Console.WriteLine("Document created, protected, and verified successfully.");
    }
}
