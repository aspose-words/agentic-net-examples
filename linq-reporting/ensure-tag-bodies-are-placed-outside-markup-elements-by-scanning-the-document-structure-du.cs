using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Sample property used in the template.
    public string Name { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Register the code page provider required by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create an output folder for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The LINQ Reporting tag is placed inside the same run with surrounding text.
        builder.Writeln("Hello <<[model.Name]>> world!");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and preprocess it.
        //    Ensure that each tag resides in its own run (outside markup).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Regular expression that matches a LINQ Reporting tag.
        Regex tagRegex = new Regex(@"(<<[^>]+>>)", RegexOptions.Compiled);

        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Iterate backwards because we will modify the Runs collection.
            for (int runIdx = paragraph.Runs.Count - 1; runIdx >= 0; runIdx--)
            {
                Run run = (Run)paragraph.Runs[runIdx];
                string text = run.Text;

                // If the run does not contain a tag, skip it.
                if (!tagRegex.IsMatch(text))
                    continue;

                // Split the run text into plain parts and tag parts.
                string[] parts = tagRegex.Split(text);

                // Remove the original run.
                paragraph.Runs.RemoveAt(runIdx);

                // Insert new runs for each non‑empty part, preserving order.
                int insertIdx = runIdx;
                foreach (string part in parts)
                {
                    if (string.IsNullOrEmpty(part))
                        continue;

                    Run newRun = (Run)run.Clone(false);
                    newRun.Text = part;
                    paragraph.Runs.Insert(insertIdx, newRun);
                    insertIdx++;
                }
            }
        }

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        Model model = new Model();

        // -----------------------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(outputDir, "output.docx");
        doc.Save(outputPath);
    }
}
