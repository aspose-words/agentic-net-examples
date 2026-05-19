using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file that will be loaded later.
        Document seedDoc = new Document();
        DocumentBuilder seedBuilder = new DocumentBuilder(seedDoc);
        seedBuilder.Writeln("Sample document for loading.");
        seedDoc.Save("input.doc");

        // Load the existing DOC file.
        Document doc = new Document("input.doc");

        // Create a date picker content control (inline SDT).
        StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline)
        {
            Title = "DatePicker",
            Tag = "date-picker"
        };

        // Configure display and storage settings.
        dateSdt.DateDisplayLocale = CultureInfo.GetCultureInfo("en-US").LCID;
        dateSdt.DateDisplayFormat = "dd MMMM, yyyy";
        dateSdt.DateStorageFormat = SdtDateStorageFormat.DateTime;
        dateSdt.CalendarType = SdtCalendarType.Gregorian;
        dateSdt.FullDate = DateTime.Today;

        // Insert the content control into the first paragraph of the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(dateSdt);

        // Save the modified document as DOCX.
        doc.Save("output.docx");
    }
}
