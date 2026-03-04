using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;

class TxtVariableExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a document variable.
        doc.Variables.Add("FullName", "John Doe");

        // Insert a DOCVARIABLE field that displays the variable's value.
        FieldDocVariable varField = (FieldDocVariable)builder.InsertField(FieldType.FieldDocVariable, true);
        varField.VariableName = "FullName";
        varField.Update(); // First evaluation shows the initial value.

        // Change the variable's value.
        doc.Variables["FullName"] = "Jane Smith";

        // Update the field so it reflects the new value.
        varField.Update();

        // Demonstrate a find-and-replace that uses a specific replacement format (Markdown).
        FindReplaceOptions replaceOptions = new FindReplaceOptions
        {
            ReplacementFormat = ReplacementFormat.Markdown
        };
        // Replace the placeholder with a markdown‑styled name.
        doc.Range.Replace("_FullName_", "**Jane Smith**", replaceOptions);

        // Configure TXT save options – set a custom paragraph break.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            ParagraphBreak = "\r\n---\r\n"
        };

        // Save the document as plain text.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DocumentWithVariables.txt");
        doc.Save(outputPath, txtOptions);

        // Output the saved text to the console for verification.
        Console.WriteLine(File.ReadAllText(outputPath));
    }
}
