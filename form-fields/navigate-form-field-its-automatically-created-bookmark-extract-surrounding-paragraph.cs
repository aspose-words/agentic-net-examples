using System;
using Aspose.Words;
using Aspose.Words.Fields;

class ExtractParagraphAroundFormField
{
    static void Main()
    {
        // Create a document with a form field if an input file is not present.
        Document doc;
        const string inputPath = "Input.docx";

        if (System.IO.File.Exists(inputPath))
        {
            doc = new Document(inputPath);
        }
        else
        {
            doc = new Document();
            DocumentBuilder tempBuilder = new DocumentBuilder(doc);
            tempBuilder.Writeln("Paragraph before the form field.");
            // Insert a checkbox form field; this automatically creates a bookmark with the same name.
            FormField formField = tempBuilder.InsertCheckBox("MyCheckBox", false, 0);
            tempBuilder.Writeln("Paragraph after the form field.");
        }

        // Ensure the document actually contains at least one form field.
        if (doc.Range.FormFields.Count == 0)
        {
            Console.WriteLine("No form fields found in the document.");
            return;
        }

        // Use the first form field.
        FormField targetField = doc.Range.FormFields[0];
        string bookmarkName = targetField.Name;

        // Navigate to the bookmark that was automatically created for the form field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        if (!builder.MoveToBookmark(bookmarkName))
        {
            Console.WriteLine($"Bookmark '{bookmarkName}' not found.");
            return;
        }

        // The current paragraph now contains the form field.
        Paragraph surroundingParagraph = builder.CurrentParagraph;
        if (surroundingParagraph == null)
        {
            Console.WriteLine("Unable to locate the surrounding paragraph.");
            return;
        }

        string paragraphText = surroundingParagraph.GetText();

        Console.WriteLine("Paragraph containing the form field:");
        Console.WriteLine(paragraphText.Trim());

        // Save the document (optional).
        doc.Save("Output.docx");
    }
}
