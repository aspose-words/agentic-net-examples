using System;
using Aspose.Words;
using Aspose.Words.Fields; // <-- added namespace for TextFormFieldType

class Program
{
    static void Main()
    {
        // Path where the document will be saved.
        string outputPath = @"C:\Temp\FormFieldsOnlyEditable.docx";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Content that should be read‑only.
        // -----------------------------------------------------------------
        builder.Writeln("This paragraph is read‑only and cannot be edited by the user.");

        // -----------------------------------------------------------------
        // Insert a form field – this will be the only editable part.
        // -----------------------------------------------------------------
        builder.Write("Enter your name: ");
        builder.InsertTextInput(
            "NameField",                     // name of the form field
            TextFormFieldType.Regular,       // regular text field
            "",                              // default value (empty)
            "John Doe",                      // placeholder text
            0);                               // maximum length (0 = unlimited)

        // Add more read‑only content after the form field.
        builder.Writeln("Another read‑only paragraph follows the form field.");

        // -----------------------------------------------------------------
        // Protect the whole document so that only form fields are editable.
        // -----------------------------------------------------------------
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Ensure every section remains protected for forms (default is true,
        // but we reset it explicitly in case any section was altered elsewhere).
        foreach (Section section in doc.Sections)
        {
            section.ProtectedForForms = true;
        }

        // -----------------------------------------------------------------
        // Save the resulting document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
