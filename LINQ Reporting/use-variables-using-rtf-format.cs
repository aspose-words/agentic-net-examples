using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add document variables.
        doc.Variables.Add("Author", "John Doe");
        doc.Variables.Add("Title", "Sample RTF Document");

        // Use DocumentBuilder to insert fields that display the variables.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document Variables (RTF format):");

        // Insert a DOCVARIABLE field for the Author variable.
        FieldDocVariable authorField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        authorField.VariableName = "Author";
        authorField.Update();

        builder.Writeln(); // Add a line break.

        // Insert a DOCVARIABLE field for the Title variable.
        FieldDocVariable titleField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        titleField.VariableName = "Title";
        titleField.Update();

        // Configure RTF save options.
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            ExportCompactSize = true,          // Reduce file size (no RTL support needed).
            ExportImagesForOldReaders = false, // Smaller file, old‑reader compatibility not required.
            PrettyFormat = false               // No extra whitespace.
        };

        // Save the document as an RTF file using the specified options.
        doc.Save("Variables.rtf", saveOptions);
    }
}
