using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple holder class used as a data source for the reporting engine.
    public class DocumentHolder
    {
        public Document Document { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains LINQ Reporting syntax.
            // -----------------------------------------------------------------
            Document template = new Document();                     // create a blank document
            DocumentBuilder builder = new DocumentBuilder(template);

            // The template uses a foreach loop over a collection named "src".
            // For each item we insert the document referenced by src.Document.
            builder.Writeln("<<foreach [src]>>");
            builder.Writeln("<<doc [src.Document]>>");
            builder.Writeln("<</foreach>>");

            // ---------------------------------------------------------------
            // 2. Load all Markdown (*.md) files from a folder into Document objects.
            // ---------------------------------------------------------------
            string markdownFolder = Path.Combine(Environment.CurrentDirectory, "InputMd");
            // Ensure the folder exists; in a real scenario handle missing folder appropriately.
            if (!Directory.Exists(markdownFolder))
                Directory.CreateDirectory(markdownFolder);

            // Load each .md file as an Aspose.Words Document (Markdown is auto‑detected).
            List<DocumentHolder> holders = Directory.GetFiles(markdownFolder, "*.md")
                .Select(filePath => new DocumentHolder { Document = new Document(filePath) })
                .ToList();

            // ---------------------------------------------------------------
            // 3. Populate the template using the ReportingEngine.
            // ---------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // The data source is the collection of DocumentHolder objects; we give it the name "src".
            engine.BuildReport(template, holders, "src");

            // ---------------------------------------------------------------
            // 4. Save the resulting document.
            // ---------------------------------------------------------------
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docx");
            template.Save(outputPath);
        }
    }
}
