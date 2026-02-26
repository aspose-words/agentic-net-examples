using System;
using Aspose.Words;
using Aspose.Words.Saving;

class MergeFieldTextInsertion
{
    static void Main()
    {
        // Paths to the source DOCX and the target PDF.
        const string inputFile = @"C:\Docs\Template.docx";
        const string outputFile = @"C:\Docs\Result.pdf";

        // Load the existing Word document.
        Document doc = new Document(inputFile);

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Insert text BEFORE the merge field -----
        // Move the cursor to the position just BEFORE the specified merge field.
        // The third parameter (isDeleteField) is false so the field remains in the document.
        if (builder.MoveToMergeField("MyMergeField", isAfter: false, isDeleteField: false))
        {
            // Write the desired text at the current cursor position.
            builder.Write("Text before the field ");
        }

        // ----- Insert text AFTER the merge field -----
        // Move the cursor to the position just AFTER the same merge field.
        if (builder.MoveToMergeField("MyMergeField", isAfter: true, isDeleteField: false))
        {
            // Write the desired text after the field.
            builder.Write(" Text after the field");
        }

        // Save the modified document as PDF.
        doc.Save(outputFile, SaveFormat.Pdf);
    }
}
