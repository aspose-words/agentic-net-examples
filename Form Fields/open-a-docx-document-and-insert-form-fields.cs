using System;
using Aspose.Words;
using Aspose.Words.Fields;

class InsertFormFieldsExample
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\Template.docx";

        // Load the document from file.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document so that new content is appended.
        builder.MoveToDocumentEnd();

        // -------------------------------------------------
        // 1. Insert a regular text input form field.
        // -------------------------------------------------
        builder.Writeln("Enter your name:");
        // Parameters: name, type, format, default text, max length.
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter name here", 50);
        builder.InsertParagraph();

        // -------------------------------------------------
        // 2. Insert a combo box (drop‑down) form field.
        // -------------------------------------------------
        builder.Writeln("Select your favorite color:");
        string[] colors = { "Red", "Green", "Blue", "Other" };
        // Parameters: name, items, selected index.
        builder.InsertComboBox("ColorField", colors, 0);
        builder.InsertParagraph();

        // -------------------------------------------------
        // 3. Insert a check box form field.
        // -------------------------------------------------
        builder.Writeln("Agree to the terms:");
        // Parameters: name, isCheckedByDefault, size (points).
        builder.InsertCheckBox("AgreeField", false, 50);
        builder.InsertParagraph();

        // -------------------------------------------------
        // Optional: protect the document so that only form fields can be edited.
        // -------------------------------------------------
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the modified document.
        string outputPath = @"C:\Docs\TemplateWithFormFields.docx";
        doc.Save(outputPath);
    }
}
