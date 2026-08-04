using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class DeleteAllFormFields
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a few different form fields.
        builder.Write("Choose a value: ");
        builder.InsertComboBox("ComboBox", new[] { "One", "Two", "Three" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        builder.Write("Check this box: ");
        builder.InsertCheckBox("CheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        builder.Write("Enter text: ");
        builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Placeholder", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document with form fields (optional, just to see the initial state).
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string initialPath = Path.Combine(outputDir, "DocumentWithFormFields.docx");
        doc.Save(initialPath);

        // Iterate over the FormFields collection and remove each form field.
        // Collect fields first to avoid modifying the collection while iterating.
        List<FormField> fieldsToRemove = new List<FormField>();
        foreach (FormField field in doc.Range.FormFields)
        {
            fieldsToRemove.Add(field);
        }

        foreach (FormField field in fieldsToRemove)
        {
            field.RemoveField();
        }

        // Save the cleaned document.
        string cleanedPath = Path.Combine(outputDir, "DocumentWithoutFormFields.docx");
        doc.Save(cleanedPath);
    }
}
