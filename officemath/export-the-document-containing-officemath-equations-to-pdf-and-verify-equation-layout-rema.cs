using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Document with OfficeMath equations:");

        // Insert first equation using the deterministic EQ‑field bootstrap workflow.
        InsertEquation(builder, @"\f(1,2)"); // simple fraction

        // Add a line break between equations.
        builder.Writeln();

        // Insert second equation.
        InsertEquation(builder, @"\r(3,x)"); // cube root

        // Ensure all fields are up‑to‑date before saving.
        doc.UpdateFields();

        // Save the document as PDF.
        string pdfPath = "OfficeMathExport.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF file at '{pdfPath}'.");

        // Verify that the original document contains the expected top‑level OfficeMath nodes.
        var topLevelEquations = doc.GetChildNodes(NodeType.OfficeMath, true)
                                   .Cast<OfficeMath>()
                                   .Where(om => om.MathObjectType == MathObjectType.OMathPara)
                                   .ToList();

        if (topLevelEquations.Count != 2)
            throw new InvalidOperationException($"Expected 2 top‑level equations, but found {topLevelEquations.Count}.");

        // If execution reaches this point, the export succeeded and the equations are present.
        Console.WriteLine("PDF exported successfully and equation layout verified.");
    }

    // Helper method that creates a real OfficeMath node from an EQ field and inserts it into the document.
    private static void InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an empty Equation field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = (FieldEQ)field;

        // Move to the field separator and write the EQ argument.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(eqSwitch);

        // Update the field so that the EQ argument is processed.
        field.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath? officeMath = fieldEQ.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        Node fieldStart = fieldEQ.Start;
        CompositeNode parent = (CompositeNode)fieldStart.ParentNode;
        parent.InsertBefore(officeMath, fieldStart);

        // Remove the original field, leaving only the OfficeMath node.
        fieldEQ.Remove();
    }
}
