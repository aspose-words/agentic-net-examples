using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AppendDocumentExample
{
    // Data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Flag that determines whether the appendix should be included.
        public bool IncludeAppendix { get; set; }

        // The appendix document to be merged when the flag is true.
        public Document? AppendixDoc { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary files.
            const string templatePath = "Template.docx";
            const string appendixPath = "Appendix.docx";
            const string outputPath = "Result.docx";

            // -------------------------------------------------
            // 1. Create the main template document.
            // -------------------------------------------------
            Document template = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(template);

            tmplBuilder.Writeln("=== Main Report ===");
            tmplBuilder.Writeln("This part of the document is always present.");

            // Conditional block: include the appendix only when IncludeAppendix == true.
            tmplBuilder.Writeln("<<if [IncludeAppendix]>>");
            // The <<doc>> tag inserts another document.
            tmplBuilder.Writeln("<<doc [AppendixDoc]>>");
            tmplBuilder.Writeln("<</if>>");

            // Save the template to disk.
            template.Save(templatePath);

            // -------------------------------------------------
            // 2. Create the appendix document that may be merged.
            // -------------------------------------------------
            Document appendix = new Document();
            DocumentBuilder appBuilder = new DocumentBuilder(appendix);
            appBuilder.Writeln("=== Appendix ===");
            appBuilder.Writeln("This content is added only when the flag is true.");
            appendix.Save(appendixPath);

            // -------------------------------------------------
            // 3. Load the documents back (simulating a real scenario).
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            Document loadedAppendix = new Document(appendixPath);

            // -------------------------------------------------
            // 4. Prepare the data model.
            // -------------------------------------------------
            ReportModel model = new ReportModel
            {
                IncludeAppendix = true,          // Change to false to omit the appendix.
                AppendixDoc = loadedAppendix
            };

            // -------------------------------------------------
            // 5. Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root name in the template is "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // -------------------------------------------------
            // 6. Save the final document.
            // -------------------------------------------------
            loadedTemplate.Save(outputPath);
        }
    }
}
