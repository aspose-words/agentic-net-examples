using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

namespace AsposeWordsDatePickerExample
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the input DOC and output DOCX files.
            const string inputPath = "Sample.doc";
            const string outputPath = "SampleWithDatePicker.docx";

            // Ensure a sample DOC file exists.
            CreateSampleDocIfMissing(inputPath);

            // Load the existing DOC file.
            Document doc = new Document(inputPath);

            // Prepare a builder to modify the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            // Create a date picker content control (StructuredDocumentTag of type Date).
            StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline)
            {
                // Optional display settings.
                DateDisplayLocale = CultureInfo.GetCultureInfo("en-US").LCID,
                DateDisplayFormat = "MMMM dd, yyyy",
                DateStorageFormat = SdtDateStorageFormat.DateTime,
                CalendarType = SdtCalendarType.Gregorian,
                FullDate = DateTime.Today
            };

            // Insert the content control into the document.
            builder.InsertNode(dateSdt);

            // Save the result as DOCX.
            doc.Save(outputPath, SaveFormat.Docx);
        }

        private static void CreateSampleDocIfMissing(string path)
        {
            if (File.Exists(path))
                return;

            // Create a simple document with one paragraph.
            Document tempDoc = new Document();
            DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
            tempBuilder.Writeln("This is a sample document created for the date picker example.");

            // Save it in the legacy DOC format.
            tempDoc.Save(path, SaveFormat.Doc);
        }
    }
}
