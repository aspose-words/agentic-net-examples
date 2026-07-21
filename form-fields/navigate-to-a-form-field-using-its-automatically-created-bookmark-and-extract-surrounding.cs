using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Path for the temporary document.
        const string filePath = "FormFieldDoc.docx";

        // 1. Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Write some surrounding text.
        builder.Writeln("This is the paragraph before the form field.");

        // 3. Insert a text input form field with a name.
        //    A bookmark with the same name ("MyField") is created automatically.
        FormField formField = builder.InsertTextInput(
            "MyField",                     // name of the form field (bookmark)
            TextFormFieldType.Regular,    // type of the field
            "",                           // format (none)
            "Default value",              // initial displayed text
            0);                           // no length limit

        // 4. Write more text after the field.
        builder.Writeln("This is the paragraph after the form field.");

        // 5. Save the document.
        doc.Save(filePath);

        // 6. Load the document back (demonstrates the load rule).
        Document loadedDoc = new Document(filePath);

        // 7. Retrieve the form field via its automatically created bookmark.
        FormField retrievedField = loadedDoc.Range.FormFields["MyField"];
        if (retrievedField == null)
        {
            throw new InvalidOperationException("Form field 'MyField' was not found.");
        }

        // 8. Get the paragraph that contains the form field.
        Paragraph parentParagraph = retrievedField.ParentParagraph;
        if (parentParagraph == null)
        {
            throw new InvalidOperationException("Parent paragraph of the form field is null.");
        }

        // 9. Extract the full text of that paragraph (including the field result).
        string paragraphText = parentParagraph.GetText();

        // 10. Output the extracted paragraph text.
        Console.WriteLine("Paragraph containing the form field:");
        Console.WriteLine(paragraphText);

        // 11. Save the document again (even if unchanged) to satisfy the save rule.
        loadedDoc.Save(filePath);
    }
}
