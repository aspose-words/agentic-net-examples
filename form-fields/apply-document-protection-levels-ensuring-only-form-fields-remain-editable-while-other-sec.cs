using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1 – regular text. After protection this part will be read‑only.
        builder.Writeln("Section 1: This text will be read‑only.");
        // Insert a page break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 – contains a form field that the user can edit.
        builder.Writeln("Section 2: Please fill the form field below.");
        builder.Write("Enter your name: ");
        FormField nameField = builder.InsertTextInput(
            "NameField",                     // field name
            TextFormFieldType.Regular,       // field type
            "",                              // default text (empty)
            "Your name",                     // placeholder text
            50);                             // maximum length

        // Set an initial value for the form field (optional).
        nameField.TextInputDefault = "John Doe";

        // Protect the whole document so that only form fields are editable.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the resulting document.
        doc.Save("FormFieldsProtected.docx");
    }
}
