using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModulesReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source document that may contain VBA macros
            Document srcDoc = new Document("InputDocument.docm");

            // Access the VBA project; if there is none, the report will be empty
            VbaProject vbaProject = srcDoc.VbaProject;

            // Create a new blank document to hold the report
            Document reportDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(reportDoc);

            // Write a header for the report
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("VBA Modules Report");

            // Switch to normal style for the table rows
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

            // Guard against documents without a VBA project
            if (vbaProject != null && vbaProject.Modules != null)
            {
                // Iterate through each VBA module in the project
                foreach (VbaModule module in vbaProject.Modules)
                {
                    // Determine the type of the module (Procedural, Document, etc.)
                    VbaModuleType moduleType = module.Type;

                    // Count the number of lines in the module's source code
                    int lineCount = 0;
                    if (!string.IsNullOrEmpty(module.SourceCode))
                    {
                        string[] lines = module.SourceCode.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                        lineCount = lines.Length;
                    }

                    // Write the module information to the report
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

            // Save the report document
            reportDoc.Save("VbaModulesReport.docx");
        }
    }
}
