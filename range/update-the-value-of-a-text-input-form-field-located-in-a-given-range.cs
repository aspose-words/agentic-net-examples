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

        // Insert a text input form field named "MyTextField" with an initial value.
        // Parameters: name, type, format, default text, maximum length (0 = no limit).
        builder.InsertTextInput("MyTextField", TextFormFieldType.Regular, "", "Initial value", 0);
        builder.Writeln(); // End the paragraph.

        // Save the original document (optional, just to show the before state).
        doc.Save("Original.docx");

        // Locate the form field by its name using the document's FormFields collection.
        FormField formField = doc.Range.FormFields["MyTextField"];
        if (formField != null)
        {
            // Update the value (result) of the text input form field.
            formField.Result = "Updated value";
        }

        // Demonstrate accessing the form field via the range that contains it.
        // The form field resides inside a paragraph; we can get that paragraph's range.
        // This shows how a range can be used to work with nodes inside it.
        // Paragraph paragraph = (Paragraph)formField.Start.ParentNode;
        // paragraph.Range.FormFields[0].Result = "Updated via range";

        // Save the modified document.
        doc.Save("Updated.docx");
    }
}
