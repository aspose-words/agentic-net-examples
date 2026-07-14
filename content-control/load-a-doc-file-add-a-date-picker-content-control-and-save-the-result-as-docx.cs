using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a simple source DOC file if it does not exist.
        const string inputPath = "input.doc";
        if (!System.IO.File.Exists(inputPath))
        {
            Document seedDoc = new Document();
            DocumentBuilder seedBuilder = new DocumentBuilder(seedDoc);
            seedBuilder.Writeln("Sample document for date picker content control.");
            seedDoc.Save(inputPath);
        }

        // Step 2: Load the existing DOC file.
        Document doc = new Document(inputPath);

        // Step 3: Create a date picker (Date) content control.
        StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline)
        {
            Title = "AppointmentDate",
            Tag = "appointment-date",
            DateDisplayFormat = "dd MMMM, yyyy",
            DateStorageFormat = SdtDateStorageFormat.DateTime,
            CalendarType = SdtCalendarType.Gregorian,
            FullDate = DateTime.Today
        };

        // Step 4: Insert the content control into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();
        builder.InsertNode(dateSdt);

        // Step 5: Save the modified document as DOCX.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
