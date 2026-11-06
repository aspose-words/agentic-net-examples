using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add introductory text.
        builder.Writeln("Please accept the terms:");

        // Insert an empty paragraph that will contain the checkbox.
        builder.Writeln();
        Paragraph checkboxParagraph = (Paragraph)builder.CurrentParagraph;

        // Move the builder to the start of the paragraph.
        builder.MoveTo(checkboxParagraph);

        // Insert a checkbox form field named "AcceptTerms" with default checked state.
        // The third parameter is the size of the checkbox (in points).
        builder.InsertCheckBox("AcceptTerms", true, 50);

        // Add text after the checkbox in the same paragraph.
        builder.Write(" I agree.");
        builder.Writeln(); // End the paragraph.

        // Save the document.
        doc.Save("CheckboxFormField.docx");
    }
}
