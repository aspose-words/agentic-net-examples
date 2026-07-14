using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Path for the macro‑enabled document.
            string docPath = Path.Combine(outputDir, "DocumentWithMacro.docm");

            // Create a new blank document and add a simple table.
            Document doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Sample table:");
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Data 1");
            builder.InsertCell();
            builder.Write("Data 2");
            builder.EndTable();

            // Ensure the document has a VBA project.
            if (doc.VbaProject == null)
            {
                VbaProject project = new VbaProject
                {
                    Name = "AsposeProject"
                };
                doc.VbaProject = project;
            }

            // Create a new VBA module with a macro that formats all tables.
            VbaModule module = new VbaModule
            {
                Name = "TableFormatter",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub AutoFormatTables()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Style = ""Table Grid""
        tbl.Rows.HeightRule = wdRowHeightExactly
        tbl.Rows.Height = InchesToPoints(0.2)
    Next tbl
End Sub"
            };

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(module);

            // Save the document in a macro‑enabled format.
            doc.Save(docPath, SaveFormat.Docm);
        }
    }
}
