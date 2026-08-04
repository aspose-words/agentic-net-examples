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
            // Prepare a folder for the sample files.
            string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            Directory.CreateDirectory(dataDir);

            // Path to the initial DOCX file.
            string docxPath = Path.Combine(dataDir, "Sample.docx");

            // Create a simple DOCX document with a table if it does not already exist.
            if (!File.Exists(docxPath))
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                // Insert a 2x2 table with sample data.
                builder.StartTable();
                builder.InsertCell();
                builder.Write("Header 1");
                builder.InsertCell();
                builder.Write("Header 2");
                builder.EndRow();

                builder.InsertCell();
                builder.Write("Cell 1");
                builder.InsertCell();
                builder.Write("Cell 2");
                builder.EndTable();

                doc.Save(docxPath);
            }

            // Load the DOCX document.
            Document loadedDoc = new Document(docxPath);

            // Ensure the document has a VBA project; create one if missing.
            if (loadedDoc.VbaProject == null)
            {
                loadedDoc.VbaProject = new VbaProject();
                loadedDoc.VbaProject.Name = "AsposeProject";
            }

            // Define a VBA macro that formats all tables in the document.
            string vbaCode = @"
Sub AutoFormatTables()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Range.Font.Name = ""Calibri""
        tbl.Range.Font.Size = 11
        tbl.Rows.HeightRule = wdRowHeightExactly
        tbl.Rows.Height = InchesToPoints(0.25)
        tbl.Borders.Enable = True
    Next tbl
End Sub
";

            // Create a new VBA module and set its properties.
            VbaModule module = new VbaModule();
            module.Name = "TableFormatter";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = vbaCode;

            // Add the module to the VBA project.
            loadedDoc.VbaProject.Modules.Add(module);

            // Save the updated document as a macro‑enabled file.
            string outputPath = Path.Combine(dataDir, "SampleWithMacro.docm");
            loadedDoc.Save(outputPath);
        }
    }
}
