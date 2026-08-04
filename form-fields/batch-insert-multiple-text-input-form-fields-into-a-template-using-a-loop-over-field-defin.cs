using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    // Simple definition for a text input form field.
    private struct TextFieldDefinition
    {
        public string Name;          // Form field name (also bookmark name)
        public string DefaultText;   // Text that appears when the field is empty
        public int MaxLength;        // Maximum number of characters (0 = unlimited)

        public TextFieldDefinition(string name, string defaultText, int maxLength)
        {
            Name = name;
            DefaultText = defaultText;
            MaxLength = maxLength;
        }
    }

    public static void Main()
    {
        // Define a set of text input fields to be inserted.
        TextFieldDefinition[] fields = new TextFieldDefinition[]
        {
            new TextFieldDefinition("FirstName", "Enter first name", 30),
            new TextFieldDefinition("LastName", "Enter last name", 30),
            new TextFieldDefinition("Email", "example@domain.com", 50),
            new TextFieldDefinition("Phone", "123-456-7890", 20)
        };

        // Create a new blank document and a builder to edit it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a title paragraph.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("User Information Form");
        builder.InsertParagraph();

        // Loop over the field definitions and insert each one.
        foreach (var def in fields)
        {
            // Write a label for the field.
            builder.Font.Size = 12;
            builder.Font.Bold = false;
            builder.Write($"{def.Name}: ");

            // Insert the text input form field.
            // Parameters: name, type, format, default text, max length.
            builder.InsertTextInput(def.Name, TextFormFieldType.Regular, "", def.DefaultText, def.MaxLength);

            // Move to the next line after each field.
            builder.InsertParagraph();
        }

        // Validate that the expected number of form fields were created.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count != fields.Length)
        {
            throw new InvalidOperationException($"Expected {fields.Length} form fields, but found {formFields.Count}.");
        }

        // Optionally, write the names of the created fields to the console.
        Console.WriteLine("Created form fields:");
        foreach (FormField field in formFields)
        {
            Console.WriteLine($"- {field.Name}");
        }

        // Save the document to disk.
        string outputPath = "BatchFormFields.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
