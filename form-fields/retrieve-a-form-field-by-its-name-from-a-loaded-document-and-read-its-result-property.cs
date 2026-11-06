using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        const string fileName = "FormFields.docx";
        const string updatedFileName = "FormFieldsUpdated.docx";
        const string fieldName = "CustomerName";

        // Create a new document and insert a text input form field
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please enter your name:");
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", "John Doe", 0);
        doc.Save(fileName);

        // Load the document
        Document loadedDoc = new Document(fileName);
        FormFieldCollection fields = loadedDoc.Range.FormFields;
        FormField field = fields[fieldName];
        if (field == null)
        {
            throw new InvalidOperationException($"Form field '{fieldName}' not found.");
        }

        // Read the Result property
        string originalResult = field.Result;
        Console.WriteLine($"Original value of '{fieldName}': {originalResult}");

        // Update the field value
        field.Result = "Jane Smith";

        // Validate the update
        if (field.Result != "Jane Smith")
        {
            throw new InvalidOperationException("Failed to update the form field value.");
        }

        // Save the updated document
        loadedDoc.Save(updatedFileName);
        Console.WriteLine($"Updated document saved as '{updatedFileName}'.");
    }
}
