using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert non‑field content.
        builder.Writeln("This is non‑field content that should not be editable after protection.");

        // Insert a text input form field.
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default text", 0);

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBoxField", false, 0);

        // Save the initial document.
        doc.Save("FormFields.docx");

        // Protect the document to allow only form field editing.
        doc.Protect(ProtectionType.AllowOnlyFormFields, "password");

        // Attempt to edit non‑field content (should fail).
        bool nonFieldEditFailed = false;
        Paragraph firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null && firstParagraph.Runs.Count > 0)
        {
            try
            {
                firstParagraph.Runs[0].Text = "Attempted edit of non‑field content.";
            }
            catch (InvalidOperationException)
            {
                nonFieldEditFailed = true;
            }
            catch (Exception)
            {
                // Any other exception also indicates edit is not allowed.
                nonFieldEditFailed = true;
            }
        }

        // Verify that editing a form field is allowed.
        bool formFieldEditSucceeded = false;
        FormField textField = doc.Range.FormFields["TextField"];
        FormField checkBoxField = doc.Range.FormFields["CheckBoxField"];
        if (textField != null && checkBoxField != null)
        {
            try
            {
                textField.Result = "Updated value";
                checkBoxField.Checked = true;
                formFieldEditSucceeded = true;
            }
            catch (Exception)
            {
                formFieldEditSucceeded = false;
            }
        }

        // Save the protected document.
        doc.Save("FormFields_Protected.docx");

        // Output verification results.
        Console.WriteLine(nonFieldEditFailed
            ? "Non‑field content could not be edited as expected."
            : "Non‑field content was edited unexpectedly.");

        Console.WriteLine(formFieldEditSucceeded
            ? "Form fields were edited successfully."
            : "Failed to edit form fields.");
    }
}
