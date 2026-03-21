using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModulesReport
{
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "InputDocument.docm";
            Document srcDoc;

            // Load the source document if it exists; otherwise create an empty document.
            if (File.Exists(inputPath))
            {
                srcDoc = new Document(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating an empty document for the report.");
                srcDoc = new Document();
            }

            // Access the VBA project; if there is none, the report will be empty.
            VbaProject vbaProject = srcDoc.VbaProject;

            // Create a new blank document to hold the report.
            Document reportDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(reportDoc);

            // Write a header for the report.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("VBA Modules Report");

            // Switch to normal style for the table rows.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

            // Guard against documents without a VBA project.
            if (vbaProject != null && vbaProject.Modules != null && vbaProject.Modules.Count > 0)
            {
                // Iterate through each VBA module in the project.
                foreach (VbaModule module in vbaProject.Modules)
                {
                    // Determine the type of the module (Procedural, Document, etc.).
                    VbaModuleType moduleType = module.Type;

                    // Count the number of lines in the module's source code.
                    int lineCount = 0;
                    if (!string.IsNullOrEmpty(module.SourceCode))
                    {
                        string[] lines = module.SourceCode.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                        lineCount = lines.Length;
                    }

                    // Write the module information to the report.
                    builder.Writeln($"Module Name: {module.Name}");
                    builder.Writeln($"Module Type: {moduleType}");
                    builder.Writeln($"Lines of Code: {lineCount}");
                    builder.Writeln(); // blank line between modules
                }
            }
            else
            {
                builder.Writeln("No VBA project found in the source document.");
            }

            // Save the report document.
            const string outputPath = "VbaModulesReport.docx";
            reportDoc.Save(outputPath);
            Console.WriteLine($"Report saved to '{outputPath}'.");
        }
    }
}
