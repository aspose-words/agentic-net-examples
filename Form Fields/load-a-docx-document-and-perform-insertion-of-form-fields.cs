using System;
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
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);

        // Insert a numeric text input form field.
        builder.Writeln("Enter your age:");
        builder.InsertTextInput("AgeField", TextFormFieldType.Number, "", "30", 3);

        // Insert a date text input form field.
        builder.Writeln("Enter today's date:");
        builder.InsertTextInput("DateField", TextFormFieldType.Date, "", DateTime.Today.ToShortDateString(), 10);

        // Insert a combo box with a list of options.
        builder.Writeln("Select your country:");
        string[] countries = { "USA", "Canada", "UK", "Australia" };
        builder.InsertComboBox("CountryField", countries, 0);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
