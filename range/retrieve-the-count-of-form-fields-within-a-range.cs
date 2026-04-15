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

        // Insert a combo box form field.
        builder.Write("Choose a value: ");
        builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Check this box: ");
        builder.InsertCheckBox("MyCheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter text: ");
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder text", 50);

        // Retrieve the count of form fields in the whole document range.
        int totalFormFields = doc.Range.FormFields.Count;

        // Retrieve the count of form fields in the first paragraph's range (if any).
        Paragraph firstParagraph = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[0];
        int paragraphFormFields = firstParagraph.Range.FormFields.Count;

        // Output the counts.
        Console.WriteLine($"Total form fields in document: {totalFormFields}");
        Console.WriteLine($"Form fields in first paragraph: {paragraphFormFields}");

        // Save the document (optional, demonstrates that the document can be persisted).
        doc.Save("FormFields.docx");
    }
}
