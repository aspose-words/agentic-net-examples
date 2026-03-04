using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the source PDF document.
        // Aspose.Words can directly load PDF files into a Document object.
        Document doc = new Document("SourceDocument.pdf");

        // Create a DocumentBuilder to modify the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph at the beginning of the document.
        // Move the cursor to the start of the first section.
        builder.MoveToDocumentStart();

        // Apply Heading 1 style to the paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

        // Write the required heading text.
        builder.Writeln("LINQ Reporting Introduction to LINQ Reporting Engine");

        // Save the modified document in DOCX format.
        doc.Save("ResultDocument.docx");
    }
}
