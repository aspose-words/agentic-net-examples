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

        // Insert several form fields for demonstration.
        builder.Write("Choose a value: ");
        builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        builder.Write("Check this box: ");
        builder.InsertCheckBox("MyCheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        builder.Write("Enter text: ");
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder text", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document that contains the form fields.
        doc.Save("FormFieldsBefore.docx");

        // Delete all form fields by iterating over the collection and removing each one.
        // Removing by index repeatedly avoids issues with collection modification during iteration.
        while (doc.Range.FormFields.Count > 0)
        {
            doc.Range.FormFields[0].RemoveField();
        }

        // Save the document after all form fields have been removed.
        doc.Save("FormFieldsAfter.docx");
    }
}
