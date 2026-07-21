using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = "Data";
        Directory.CreateDirectory(dataDir);

        // Paths for the intermediate DOCX and the final DOCM.
        string docxPath = Path.Combine(dataDir, "Sample.docx");
        string docmPath = Path.Combine(dataDir, "SampleWithMacro.docm");

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document that contains a table.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 2x2 table with sample text.
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

        // Save the document as DOCX.
        doc.Save(docxPath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX file we just created.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docxPath);

        // -----------------------------------------------------------------
        // 3. Ensure the document has a VBA project; create one if missing.
        // -----------------------------------------------------------------
        if (loadedDoc.VbaProject == null)
        {
            VbaProject vbaProject = new VbaProject();
            vbaProject.Name = "AsposeProject";
            loadedDoc.VbaProject = vbaProject;
        }

        // -----------------------------------------------------------------
        // 4. Add a new VBA module that formats all tables in the document.
        // -----------------------------------------------------------------
        VbaModule module = new VbaModule();
        module.Name = "TableFormatter";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub AutoFormatTables()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Range.Font.Name = ""Arial""
        tbl.Range.Font.Size = 10
        tbl.Rows.HeightRule = wdRowHeightExactly
        tbl.Rows.Height = 15
        tbl.Borders.Enable = True
    Next tbl
End Sub
";

        // Add the module to the VBA project.
        loadedDoc.VbaProject.Modules.Add(module);

        // -----------------------------------------------------------------
        // 5. Save the updated document as a macro‑enabled DOCM file.
        // -----------------------------------------------------------------
        loadedDoc.Save(docmPath);
    }
}
