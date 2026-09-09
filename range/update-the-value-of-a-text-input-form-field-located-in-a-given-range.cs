using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // ---------- Create a sample document ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory line.
        builder.Writeln("Please fill the form below:");

        // Insert a text input form field named "MyField" with default text "Old value".
        // The overload requires a format string (empty) and a maximum length (0 = no limit).
        FormField textField = builder.InsertTextInput("MyField", TextFormFieldType.Regular, "Old value", "", 0);

        // Save the initial document.
        doc.Save("Original.docx");

        // ---------- Load the document and modify the form field ----------
        Document loadedDoc = new Document("Original.docx");

        // The form field resides in the second paragraph (index 1).
        Paragraph paragraphWithField = loadedDoc.FirstSection.Body.Paragraphs[1];

        // Obtain the range of that paragraph. Use the fully qualified Aspose.Words.Range to avoid ambiguity with System.Range.
        Aspose.Words.Range targetRange = paragraphWithField.Range;

        // Update the value of the text input form field within this range.
        if (targetRange.FormFields.Count > 0)
        {
            // Apply the new value to the first form field in the range.
            targetRange.FormFields[0].SetTextInputValue("New value");
        }

        // Save the updated document.
        loadedDoc.Save("Updated.docx");
    }
}
