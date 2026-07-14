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

        // First section – regular text that will be read‑only.
        builder.Writeln("Section 1: This text is read‑only.");

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section – contains form fields that the user can edit.
        builder.Writeln("Section 2: Please fill in the form below.");

        // Text input form field.
        builder.Write("Enter name: ");
        FormField nameField = builder.InsertTextInput(
            "NameField",                     // field name
            TextFormFieldType.Regular,       // type of text field
            "",                              // default text (empty)
            "Your name",                     // placeholder text
            50);                             // maximum length

        // Check box form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Accept terms: ");
        FormField acceptCheck = builder.InsertCheckBox(
            "AcceptTerms",   // field name
            false,           // default unchecked
            50);              // size in points

        // Drop‑down (combo box) form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Choose option: ");
        FormField optionCombo = builder.InsertComboBox(
            "Options",                               // field name
            new[] { "Option1", "Option2", "Option3" }, // items
            0);                                      // default selected index

        // Protect the whole document so that only form fields are editable.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the resulting document.
        doc.Save("FormFieldsProtection.docx");
    }
}
