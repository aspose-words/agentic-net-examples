using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the checkbox.
        builder.Writeln("Please accept the terms:");

        // Insert an empty paragraph to serve as the target range.
        builder.Writeln(); // creates a new empty paragraph
        Paragraph targetParagraph = builder.CurrentParagraph;

        // Move the builder to the start of the target paragraph.
        builder.MoveTo(targetParagraph);

        // Insert a checkbox form field inside this paragraph and set it checked by default.
        // The third argument is the size of the checkbox (in points).
        builder.InsertCheckBox("AcceptTerms", true, 50);

        // Add some descriptive text after the checkbox.
        builder.Writeln(" I accept the terms and conditions.");

        // Save the document.
        doc.Save("CheckboxFormField.docx");
    }
}
