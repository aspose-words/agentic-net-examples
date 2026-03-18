using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsImageTagExample
{
    // Simple data source class that will be referenced from the template.
    public class ReportData
    {
        // The expression used in the <<image [Expression]>> tag should resolve to this property.
        public string ImagePath { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load the DOCX template that contains the <<image [Expression]>> tag.
            //    Example tag inside the template: <<image [data.ImagePath]>>
            // -----------------------------------------------------------------
            Document template = new Document("Template.docx");

            // -----------------------------------------------------------------
            // 2. Prepare the data source.
            //    The property name used in the tag (ImagePath) must match exactly.
            // -----------------------------------------------------------------
            var data = new ReportData
            {
                // Provide a full or relative path to the image that should be inserted.
                ImagePath = @"Images\SampleImage.jpg"
            };

            // -----------------------------------------------------------------
            // 3. Create the ReportingEngine and build the report.
            //    The third argument ("data") is the name used to reference the
            //    data source inside the template tags.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "data");

            // -----------------------------------------------------------------
            // 4. Save the populated document.
            // -----------------------------------------------------------------
            template.Save("Result.docx");
        }
    }
}
