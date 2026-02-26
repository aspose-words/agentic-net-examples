using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class DynamicDocumentInserter
    {
        // Sample data source class used by the reporting engine.
        public class ReportData
        {
            public string Title { get; set; }
            public DateTime GeneratedOn { get; set; }
            // Add other properties that are referenced in the DOCX template.
        }

        public static void Main()
        {
            // 1. Load the DOCX template that contains the reporting placeholders.
            //    The template can contain any LINQ Reporting syntax, e.g. <<[ds.Title]>>.
            Document template = new Document("Template.docx");

            // 2. Prepare a list of source documents that will be inserted dynamically.
            //    In a real scenario these could be discovered at runtime (e.g. from a folder or a database).
            List<string> sourceDocPaths = new List<string>
            {
                "SectionA.docx",
                "SectionB.docx",
                "SectionC.docx"
            };

            // 3. Use DocumentBuilder to position the cursor where the inserts should occur.
            //    For this example we append everything to the end of the template.
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.MoveToDocumentEnd();

            // 4. Insert each source document using the "sourceStyles" switch.
            //    The switch is implemented by enabling SmartStyleBehavior in ImportFormatOptions.
            //    This makes the engine resolve style name clashes by converting the source style
            //    into direct formatting instead of creating duplicate style definitions.
            foreach (string srcPath in sourceDocPaths)
            {
                Document srcDoc = new Document(srcPath);

                // ImportFormatOptions lives directly under the Aspose.Words namespace in recent versions.
                ImportFormatOptions importOptions = new ImportFormatOptions
                {
                    // Enable the source‑styles behaviour.
                    SmartStyleBehavior = true
                };

                // Keep the original formatting of the source document while applying the
                // SmartStyleBehavior option defined above.
                builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting, importOptions);
            }

            // 5. Build the final report using LINQ Reporting.
            //    The data source can be any POCO, DataSet, etc. Here we use a simple object.
            ReportData data = new ReportData
            {
                Title = "Combined Report",
                GeneratedOn = DateTime.Now
            };

            ReportingEngine engine = new ReportingEngine();
            // The template can reference the data source via the name "ds".
            engine.BuildReport(template, data, "ds");

            // 6. Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
