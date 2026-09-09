using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

namespace AsposeWordsDatePickerExample
{
    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a simple source DOC file if it does not exist.
            const string sourcePath = "sample.doc";
            if (!System.IO.File.Exists(sourcePath))
            {
                Document seedDoc = new Document();
                DocumentBuilder seedBuilder = new DocumentBuilder(seedDoc);
                seedBuilder.Writeln("This is a sample document.");
                seedDoc.Save(sourcePath);
            }

            // Step 2: Load the existing DOC file.
            Document doc = new Document(sourcePath);

            // Step 3: Create a Date picker content control (structured document tag).
            StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline)
            {
                Title = "AppointmentDate",
                Tag = "appointment-date",
                DateDisplayFormat = "dd MMMM, yyyy",
                DateStorageFormat = SdtDateStorageFormat.DateTime,
                CalendarType = SdtCalendarType.Gregorian,
                FullDate = DateTime.Today
            };

            // Insert the date picker into the first paragraph of the document.
            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
            firstParagraph.AppendChild(dateSdt);

            // Step 4: Save the modified document as DOCX.
            const string outputPath = "output.docx";
            doc.Save(outputPath);
        }
    }
}
