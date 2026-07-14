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

        // Insert a few different form fields.
        builder.Write("Choose a value: ");
        builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        builder.Write("Check this box: ");
        builder.InsertCheckBox("MyCheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        builder.Write("Enter text: ");
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document with form fields.
        const string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Load the document back (demonstrating load lifecycle).
        Document loadedDoc = new Document(originalPath);

        // Iterate over the form fields collection in reverse order and remove each field.
        FormFieldCollection formFields = loadedDoc.Range.FormFields;
        for (int i = formFields.Count - 1; i >= 0; i--)
        {
            formFields[i].RemoveField();
        }

        // Save the cleaned document.
        const string cleanedPath = "NoFormFields.docx";
        loadedDoc.Save(cleanedPath);
    }
}
