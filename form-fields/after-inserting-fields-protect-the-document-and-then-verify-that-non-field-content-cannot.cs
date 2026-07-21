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
        builder.Writeln("This is a regular paragraph that should not be editable after protection.");

        // Insert a text input form field.
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default value", 0);
        builder.Writeln(); // Move to next line.

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBox", false, 0);
        builder.Writeln(); // Move to next line.

        // Save the initial document.
        string initialPath = "FormFields.docx";
        doc.Save(initialPath);

        // Protect the document to allow only form field editing.
        doc.Protect(ProtectionType.AllowOnlyFormFields, "password");

        // Save the protected document.
        string protectedPath = "FormFields_Protected.docx";
        doc.Save(protectedPath);

        // Attempt to edit non‑field content (should fail).
        bool nonFieldEditSucceeded = false;
        try
        {
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];
            if (firstParagraph != null)
            {
                firstParagraph.Runs.Clear();
                firstParagraph.AppendChild(new Run(doc, "Attempted edit of non‑field content."));
                // Force a save to trigger protection enforcement.
                doc.Save("Temp.docx");
                nonFieldEditSucceeded = true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Non‑field edit blocked as expected: " + ex.Message);
        }

        if (nonFieldEditSucceeded)
        {
            Console.WriteLine("Error: Non‑field content was edited despite protection.");
        }

        // Attempt to edit a form field (should succeed).
        FormField textField = doc.Range.FormFields["TextField"];
        if (textField != null)
        {
            textField.Result = "New value";
            Console.WriteLine("Form field 'TextField' updated successfully.");
        }
        else
        {
            Console.WriteLine("Form field 'TextField' not found.");
        }

        // Save the final document after attempted edits.
        string finalPath = "FormFields_Final.docx";
        doc.Save(finalPath);
    }
}
