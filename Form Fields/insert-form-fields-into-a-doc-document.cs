using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a checkbox form field.
        // Parameters: name, defaultChecked, size (in points).
        builder.InsertCheckBox("AcceptTerms", false, 50);

        // Insert a line break after the checkbox.
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box form field.
        // Parameters: name, list of items, selected index.
        string[] footwear = { "-- Select footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other" };
        builder.InsertComboBox("FootwearChoice", footwear, 0);

        // Insert another line break.
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        // Parameters: name, type, format, default text, max length (0 = unlimited).
        builder.InsertTextInput("UserName", TextFormFieldType.Regular, "", "Enter your name here", 0);

        // Save the document to disk.
        doc.Save("FormFieldsDocument.docx");
    }
}
