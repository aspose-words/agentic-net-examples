using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("CheckBox1", false, 50);
        builder.InsertParagraph();

        // Insert a combo box form field.
        builder.Write("Select option: ");
        string[] items = { "Option A", "Option B", "Option C" };
        builder.InsertComboBox("ComboBox1", items, 0);
        builder.InsertParagraph();

        // Insert a text input form field.
        builder.Write("Enter name: ");
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Your name", 30);
        builder.InsertParagraph();

        // Ensure field results are up‑to‑date.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
