using System;
using Aspose.Words;
using Aspose.Words.Markup;

class RepeatingSectionExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a repeating section content control at the block level.
        // This control will be used by the ReportingEngine to repeat its contents.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(
            doc,
            SdtType.RepeatingSection,   // The control type that repeats.
            MarkupLevel.Block);         // Block‑level so it can contain paragraphs/tables.

        // Insert the content control into the document.
        builder.InsertNode(repeatingSection);

        // Create a paragraph that will hold the ReportingEngine tags.
        Paragraph innerParagraph = new Paragraph(doc);
        repeatingSection.AppendChild(innerParagraph);

        // Move the builder cursor into the newly created paragraph.
        builder.MoveTo(innerParagraph);

        // Write the ReportingEngine foreach syntax.
        // The engine will iterate over the "Employees" collection and insert each employee's Name.
        builder.Writeln("<<foreach [in Employees]>>");
        builder.Writeln("Name: <<[Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the document.
        doc.Save("RepeatingSectionEmployee.docx");
    }
}
