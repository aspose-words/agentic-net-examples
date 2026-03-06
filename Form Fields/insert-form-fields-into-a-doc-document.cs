using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 0);
        builder.Writeln();

        // Insert a combo box form field.
        builder.Write("Select your country: ");
        string[] countries = { "USA", "Canada", "UK", "Australia" };
        builder.InsertComboBox("CountryField", countries, 0);
        builder.Writeln();

        // Insert a check box form field.
        builder.Write("Accept terms and conditions: ");
        builder.InsertCheckBox("AcceptTerms", false, 50);
        builder.Writeln();

        // Update fields to ensure results are calculated.
        doc.UpdateFields();

        // Save the document.
        doc.Save("FormFields.docx");
    }
}
