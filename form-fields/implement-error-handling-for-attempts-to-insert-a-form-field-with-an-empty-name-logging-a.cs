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

        // Attempt to insert a text input form field with an empty name.
        string emptyName = string.Empty;
        if (string.IsNullOrEmpty(emptyName))
        {
            Console.WriteLine("Warning: Cannot insert a form field with an empty name.");
        }
        else
        {
            // This block will not be executed because the name is empty.
            builder.InsertTextInput(emptyName, TextFormFieldType.Regular, "", "Placeholder", 50);
        }

        // Insert a valid text input form field.
        string textFieldName = "UserName";
        builder.Write("Enter your name: ");
        builder.InsertTextInput(textFieldName, TextFormFieldType.Regular, "", "John Doe", 50);

        // Attempt to insert a checkbox form field with an empty name.
        if (string.IsNullOrEmpty(emptyName))
        {
            Console.WriteLine("Warning: Cannot insert a checkbox form field with an empty name.");
        }
        else
        {
            // This block will not be executed because the name is empty.
            builder.InsertCheckBox(emptyName, false, 20);
        }

        // Insert a valid checkbox form field.
        string checkBoxName = "AgreeTerms";
        builder.Write(" I agree to the terms.");
        builder.InsertCheckBox(checkBoxName, false, 20);

        // Insert a dropdown (combo box) with a valid name.
        string comboBoxName = "Country";
        builder.Write("Select country: ");
        string[] items = { "USA", "Canada", "Mexico" };
        builder.InsertComboBox(comboBoxName, items, 0);

        // Save the document to the file system.
        string outputPath = "FormFields_Output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
