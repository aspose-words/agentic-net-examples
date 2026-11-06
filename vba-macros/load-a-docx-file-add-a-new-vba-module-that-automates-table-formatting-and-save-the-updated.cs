using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Vba;

namespace VbaMacroExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a blank document.
            Document doc = new Document();

            // Ensure the document has a VBA project; create one if missing.
            if (doc.VbaProject == null)
            {
                doc.VbaProject = new VbaProject();
            }

            // VBA macro that formats all tables in the document.
            string vbaCode = @"Sub FormatTables()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Rows.HeightRule = wdRowHeightExactly
        tbl.Rows.Height = 15
    Next tbl
End Sub";

            // Create a new VBA module and set its name and source code.
            VbaModule vbaModule = new VbaModule
            {
                Name = "TableFormatter",
                SourceCode = vbaCode
            };

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(vbaModule);

            // Save the updated document as a macro‑enabled file.
            const string outputPath = "output.docm";
            doc.Save(outputPath, SaveFormat.Docm);
        }
    }
}
