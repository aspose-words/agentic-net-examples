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

        // First equation: a simple fraction \f(1,2)
        builder.Writeln("First equation:");
        InsertEquation(builder, @"\f(1,2)");

        // Second equation: a simple root \r(3,x)
        builder.Writeln("Second equation:");
        InsertEquation(builder, @"\r(3,x)");

        // Save the document as PDF.
        string pdfPath = "Equations.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF file was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create PDF file at '{pdfPath}'.");

        // Validate that the document contains the expected number of top‑level OfficeMath nodes.
        int topLevelCount = doc.GetChildNodes(NodeType.OfficeMath, true)
                               .Cast<OfficeMath>()
                               .Count(om => om.MathObjectType == MathObjectType.OMathPara);

        if (topLevelCount != 2)
            throw new InvalidOperationException($"Expected 2 top‑level equations, but found {topLevelCount}.");

        // Indicate successful execution.
        Console.WriteLine("Document exported to PDF and equation layout verified successfully.");
    }

    // Helper method that creates an OfficeMath equation using the deterministic EQ‑field bootstrap workflow.
    private static void InsertEquation(DocumentBuilder builder, string eqSwitch)
    {
        // Insert an empty Equation field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Move to the field separator and write the EQ argument.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(eqSwitch);

        // Update the field so that Aspose.Words converts it to an OfficeMath object.
        field.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start and remove the original field.
            fieldEq.Start.ParentNode.InsertBefore(officeMath, fieldEq.Start);
            fieldEq.Remove();
        }
        else
        {
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");
        }
    }
}
