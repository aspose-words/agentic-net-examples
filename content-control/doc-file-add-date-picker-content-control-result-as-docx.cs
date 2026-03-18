using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Path to the existing DOC file.
        const string inputPath = "input.doc";

        // Path where the resulting DOCX will be saved.
        const string outputPath = "output.docx";

        // Load the DOC file. This uses the Document(string) constructor rule.
        Document doc = new Document(inputPath);

        // Create a date picker content control (Structured Document Tag of type Date).
        StructuredDocumentTag dateTag = new StructuredDocumentTag(
            doc,
            SdtType.Date,          // Specify that this SDT is a date picker.
            MarkupLevel.Inline);   // Insert it inline with the surrounding text.

        // Optional: set locale, display format, storage format, and calendar type.
        dateTag.DateDisplayLocale = CultureInfo.GetCultureInfo("en-US").LCID;
        dateTag.DateDisplayFormat = "dd MMMM, yyyy";
        dateTag.DateStorageFormat = SdtDateStorageFormat.DateTime;
        dateTag.CalendarType = SdtCalendarType.Gregorian;

        // Set a default date that will be shown before the user picks a date.
        dateTag.FullDate = DateTime.Today;

        // Insert the date picker into the document using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertNode(dateTag);

        // Save the modified document as DOCX. This uses the Document.Save(string, SaveFormat) rule.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
