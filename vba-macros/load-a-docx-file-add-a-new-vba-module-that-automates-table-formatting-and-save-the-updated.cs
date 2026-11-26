using System;
using Aspose.Words;
using Aspose.Words.Vba;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a simple table.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with a table.");
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
            doc.VbaProject = new VbaProject();
        }

        // VBA macro that formats all tables in the document.
        string vbaCode = @"
Option Explicit
Sub FormatTables()
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        tbl.Rows.HeightRule = wdRowHeightExactly
        tbl.Rows.Height = 15
    Next tbl
End Sub
";

        // Create a new VBA module and set its name and source code.
        VbaModule vbaModule = new VbaModule
        {
            Name = "TableFormatter",
            SourceCode = vbaCode
        };

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document as a macro‑enabled file.
        string outputPath = "output.docm";
        doc.Save(outputPath, SaveFormat.Docm);

        // Validate that the module was added.
        bool hasModule = false;
        foreach (VbaModule mod in doc.VbaProject.Modules)
        {
            if (mod.Name == "TableFormatter")
            {
                hasModule = true;
                break;
            }
        }

        Console.WriteLine(hasModule ? "VBA module added successfully." : "Failed to add VBA module.");
    }
}
