using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string docxPath = Path.Combine(Environment.CurrentDirectory, "Sample.docx");
        string docmPath = Path.Combine(Environment.CurrentDirectory, "SampleWithMacro.docm");

        // Create a sample DOCX file with a simple table if it does not already exist.
        if (!File.Exists(docxPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a 2x2 table.
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

            doc.Save(docxPath);
        }

        // Load the DOCX file.
        Document loadedDoc = new Document(docxPath);

        // Ensure the document has a VBA project.
        if (loadedDoc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "AsposeProject";
            loadedDoc.VbaProject = project;
        }

        // Create a new VBA module that formats all tables in the document.
        VbaModule module = new VbaModule();
        module.Name = "TableFormatter";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
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

        // Add the module to the VBA project.
        loadedDoc.VbaProject.Modules.Add(module);

        // Save the document as a macro-enabled .docm file.
        loadedDoc.Save(docmPath);
    }
}
