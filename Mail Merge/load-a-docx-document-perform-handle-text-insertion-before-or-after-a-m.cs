using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Insert text BEFORE the merge field -----
        // Move the cursor to the position BEFORE the specified merge field.
        // Parameters: field name, isAfter = false (before field), isDeleteField = false (keep field).
        if (builder.MoveToMergeField("MyMergeField", false, false))
        {
            // Write the desired text at this position.
            builder.Write("Text before ");
        }

        // ----- Insert text AFTER the merge field -----
        // Move the cursor to the position AFTER the same merge field.
        // Parameters: field name, isAfter = true (after field), isDeleteField = false (keep field).
        if (builder.MoveToMergeField("MyMergeField", true, false))
        {
            // Write the desired text at this position.
            builder.Write(" Text after");
        }

        // Save the modified document as PDF.
        doc.Save("Output.pdf", SaveFormat.Pdf);
    }
}
