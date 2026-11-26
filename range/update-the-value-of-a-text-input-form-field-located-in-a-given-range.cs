using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document containing a text input form field.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a text input form field named "MyField".
        builder.Writeln("Please fill the form field below:");
        builder.InsertTextInput("MyField", TextFormFieldType.Regular, "", "Default value", 0);
        builder.Writeln(); // add a line break

        // Save the source document.
        string sourcePath = Path.Combine(outputDir, "source.docx");
        doc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and locate the form field within a specific range.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Assume the form field is inside the first paragraph after the introductory line.
        Paragraph targetParagraph = (Paragraph)loadedDoc.GetChild(NodeType.Paragraph, 1, true);
        Aspose.Words.Range targetRange = targetParagraph.Range;

        // Find the first form field that resides in the target range.
        FormField targetField = targetRange.FormFields.FirstOrDefault();

        // Update the value of the located text input form field.
        if (targetField != null)
        {
            targetField.Result = "Updated value";
        }

        // -----------------------------------------------------------------
        // 3. Save the updated document.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "updated.docx");
        loadedDoc.Save(resultPath);

        // Output verification to console (no user interaction required).
        Console.WriteLine($"Form field \"{targetField?.Name}\" updated to: {targetField?.Result}");
        Console.WriteLine($"Updated document saved to: {resultPath}");
    }
}
