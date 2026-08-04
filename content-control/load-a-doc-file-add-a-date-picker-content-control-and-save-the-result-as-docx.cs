using System;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json; // Included as required package

namespace AsposeWordsDatePickerExample
{
    public class Program
    {
        public static void Main()
        {
            // Define file names
            const string inputPath = "input.doc";
            const string outputPath = "output.docx";

            // -----------------------------------------------------------------
            // Step 1: Create a simple source DOC file if it does not already exist.
            // -----------------------------------------------------------------
            if (!System.IO.File.Exists(inputPath))
            {
                Document seedDoc = new Document();
                DocumentBuilder seedBuilder = new DocumentBuilder(seedDoc);
                seedBuilder.Writeln("This is a sample document.");
                seedDoc.Save(inputPath);
            }

            // -----------------------------------------------------------------
            // Step 2: Load the existing DOC file.
            // -----------------------------------------------------------------
            Document doc = new Document(inputPath);

            // -----------------------------------------------------------------
            // Step 3: Create a Date Picker content control (date SDT).
            // -----------------------------------------------------------------
            StructuredDocumentTag dateSdt = new StructuredDocumentTag(doc, SdtType.Date, MarkupLevel.Inline)
            {
                Title = "AppointmentDate",
                Tag = "appointment-date",
                DateDisplayFormat = "dd MMMM, yyyy",
                DateStorageFormat = SdtDateStorageFormat.DateTime,
                FullDate = DateTime.Today
            };

            // -----------------------------------------------------------------
            // Step 4: Insert the date picker into the document.
            // -----------------------------------------------------------------
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd(); // Position cursor at the end of the document.
            builder.InsertNode(dateSdt); // Insert the content control.

            // -----------------------------------------------------------------
            // Step 5: Save the modified document as DOCX.
            // -----------------------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
