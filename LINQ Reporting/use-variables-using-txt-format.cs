using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a document variable that can be referenced later.
        // Variable names are case‑insensitive.
        doc.Variables.Add("FullName", "John Doe");

        // Insert a DOCVARIABLE field that will display the value of the variable.
        // The InsertField method returns a Field object; cast it to FieldDocVariable.
        FieldDocVariable field = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        field.VariableName = "FullName";
        field.Update(); // Populate the field result with the variable's current value.

        // Add a line break after the field for readability.
        builder.Writeln();

        // Configure TXT save options.
        // Here we set a custom paragraph break string that will be inserted after each paragraph.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            ParagraphBreak = "\r\n---\r\n"
        };

        // Save the document as plain text using the configured options.
        doc.Save("Output.txt", txtOptions);
    }
}
