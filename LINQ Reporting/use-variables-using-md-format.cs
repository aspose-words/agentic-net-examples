using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Access the document's variable collection.
        VariableCollection variables = doc.Variables;

        // Add some variables that we will later display in the document.
        variables.Add("Title", "Aspose.Words Demo");
        variables.Add("Author", "John Doe");
        variables.Add("Date", DateTime.Today.ToString("yyyy-MM-dd"));

        // Use a DocumentBuilder to insert a DOCVARIABLE field for each variable.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading that will be replaced by the "Title" variable.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        FieldDocVariable titleField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        titleField.VariableName = "Title";
        titleField.Update();

        // Insert a paragraph with the "Author" variable.
        builder.Writeln();
        builder.Font.Bold = true;
        builder.Write("Author: ");
        builder.Font.Bold = false;
        FieldDocVariable authorField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        authorField.VariableName = "Author";
        authorField.Update();

        // Insert a paragraph with the "Date" variable.
        builder.Writeln();
        builder.Font.Bold = true;
        builder.Write("Date: ");
        builder.Font.Bold = false;
        FieldDocVariable dateField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        dateField.VariableName = "Date";
        dateField.Update();

        // Prepare Markdown save options.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Export underline formatting as "++" (optional demonstration).
            ExportUnderlineFormatting = true,
            // Ensure that any DOCVARIABLE fields are interpreted as Markdown text.
            // The ReplacementFormat enum is used when performing a Find/Replace operation,
            // but here we simply rely on the field result, which is plain text.
        };

        // Save the document as a Markdown file.
        doc.Save("DocumentWithVariables.md", saveOptions);
    }
}
