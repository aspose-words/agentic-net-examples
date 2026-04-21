using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;
using Aspose.Words.Tables;   // Required for the Table class

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a sample DOCX file (since external files are not assumed).
        // -----------------------------------------------------------------
        string docxPath = Path.Combine(outputDir, "Sample.docx");
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Sample document with a table.");

        // Insert a simple table to demonstrate that the macro can work on it.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        builder.EndTable();

        sampleDoc.Save(docxPath);

        // -----------------------------------------------------------------
        // Step 2: Load the DOCX file.
        // -----------------------------------------------------------------
        Document doc = new Document(docxPath);

        // -----------------------------------------------------------------
        // Step 3: Ensure the document has a VBA project.
        // -----------------------------------------------------------------
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "AsposeProject";
            doc.VbaProject = project;
        }

        // -----------------------------------------------------------------
        // Step 4: Create a new VBA module that formats all tables.
        // -----------------------------------------------------------------
        VbaModule module = new VbaModule
        {
            Name = "TableFormatter",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub FormatAllTables()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        ' AutoFit the table to its contents.
        tbl.AutoFitBehavior (wdAutoFitContent)
        ' Apply a simple style.
        tbl.Style = ""Table Grid""
    Next tbl
End Sub
"
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // -----------------------------------------------------------------
        // Step 5: Save the updated document as a macro‑enabled file.
        // -----------------------------------------------------------------
        string docmPath = Path.Combine(outputDir, "SampleWithMacro.docm");
        doc.Save(docmPath);
    }
}
