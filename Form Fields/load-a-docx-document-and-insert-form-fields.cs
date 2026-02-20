using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a regular text input form field.
        builder.Writeln("Enter your name:");
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter name", 50);

        // Insert a combo box form field.
        builder.Writeln();
        string[] footwearOptions = { "-- Select footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other" };
        builder.Writeln("Choose footwear:");
        builder.InsertComboBox("FootwearCombo", footwearOptions, 0);

        // Insert a check box form field.
        builder.Writeln();
        builder.Writeln("Accept terms:");
        builder.InsertCheckBox("AcceptTerms", false, 50);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
