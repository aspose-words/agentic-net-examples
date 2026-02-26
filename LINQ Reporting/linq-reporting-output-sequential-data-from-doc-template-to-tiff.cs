using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqToTiff
{
    // Simple data model used by the LINQ reporting engine.
    public class Person
    {
        // Initialise to avoid CS8618 warning (or make it nullable: string?).
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the DOCX template that contains LINQ reporting tags.
            //    The template should use the syntax <<foreach [ds.Persons]>><<[Name]>> <<[Age]>> <</foreach>>.
            Document doc = new Document("Template.docx");

            // 2. Prepare the data source.
            var data = new
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob",   Age = 45 },
                    new Person { Name = "Carol", Age = 27 }
                }
            };

            // 3. Build the report using the LINQ ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "ds" matches the name used in the template tags.
            engine.BuildReport(doc, data, "ds");

            // 4. Configure image save options to render each page as a separate frame in a multi‑page TIFF.
            //    In recent Aspose.Words versions the MultiPageLayout property may not be present.
            //    By default, when saving to TIFF all pages are stored as separate frames, so we can omit it.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
            // Optional: set resolution or image size if required.
            // saveOptions.Resolution = 300;
            // saveOptions.ImageSize = new System.Drawing.Size(2480, 3508); // A4 at 300 DPI

            // 5. Save the populated document as a multi‑page TIFF file.
            doc.Save("ReportOutput.tiff", saveOptions);
        }
    }
}
