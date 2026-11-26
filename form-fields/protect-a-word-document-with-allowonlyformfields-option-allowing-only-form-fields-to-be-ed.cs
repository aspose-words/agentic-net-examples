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

        // Insert a text input form field.
        builder.Writeln("Please enter your name:");
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 0);
        builder.Writeln();

        // Insert a checkbox form field.
        // The third parameter is the size of the checkbox (in points). Use 0 for default size.
        builder.Writeln("Subscribe to newsletter:");
        builder.InsertCheckBox("SubscribeField", true, 0);
        builder.Writeln();

        // Insert a dropdown (combo box) form field with items.
        builder.Writeln("Select your country:");
        builder.InsertComboBox("CountryField", new string[] { "USA", "Canada", "United Kingdom" }, 0);
        builder.Writeln();

        // Save the document before protection (optional).
        doc.Save("FormFields.docx");

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields, "myPassword");

        // Save the protected document.
        doc.Save("FormFields_Protected.docx");
    }
}
