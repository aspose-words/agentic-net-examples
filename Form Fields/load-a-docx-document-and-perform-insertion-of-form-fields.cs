using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        // This uses the Document(string) constructor – the required "load" rule.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the end of the document (or any desired position).
        builder.MoveToDocumentEnd();

        // Insert a paragraph break before adding form fields.
        builder.InsertParagraph();

        // -------------------------------------------------
        // Insert a checkbox form field.
        // Parameters: name, default checked state, size (points).
        // This follows the InsertCheckBox method signature.
        builder.InsertCheckBox("AcceptTerms", false, 50);
        builder.Writeln(" I accept the terms and conditions.");

        // Insert a combo box (drop‑down) form field.
        // Parameters: name, list of items, default selected index.
        string[] colors = { "Red", "Green", "Blue", "Yellow" };
        builder.InsertParagraph();
        builder.Write("Choose a color: ");
        builder.InsertComboBox("ColorChoice", colors, 0);
        builder.Writeln();

        // Insert a text input form field.
        // Parameters: name, type, default text, placeholder, max length.
        builder.InsertParagraph();
        builder.Write("Enter your name: ");
        builder.InsertTextInput("UserName", TextFormFieldType.Regular, "", "John Doe", 40);
        builder.Writeln();

        // Update all fields so that the form field results are current.
        // This uses the Document.UpdateFields method – the required "save" rule will be applied later.
        doc.UpdateFields();

        // Save the modified document.
        // This uses the Document.Save(string) method – the required "save" rule.
        doc.Save("Output.docx");
    }
}
