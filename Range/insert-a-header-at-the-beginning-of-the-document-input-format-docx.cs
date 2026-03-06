using System;
using Aspose.Words;

class InsertHeaderExample
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = "input.docx";

        // Path where the modified document will be saved.
        string outputPath = "output.docx";

        // Load the existing document (lifecycle: load).
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert the desired header text (adds a paragraph break after the text).
        builder.Writeln("My Header Text");

        // Save the modified document (lifecycle: save).
        doc.Save(outputPath);
    }
}
