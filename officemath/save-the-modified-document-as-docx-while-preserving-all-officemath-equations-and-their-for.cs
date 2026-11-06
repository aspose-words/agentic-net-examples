using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathExample
{
    public static void Main()
    {
        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First equation -------------------------------------------------------
        builder.Writeln("First equation:");

        // Insert an EQ field (FieldEquation) and write a simple fraction argument.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Move to the field separator and write the EQ argument.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(@"\f(1,2)"); // Fraction 1/2

        // Update the field so that Aspose.Words can convert it to OfficeMath.
        field.Update();

        // Convert the field to a real OfficeMath object.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        Node fieldStart = fieldEq.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
        fieldEq.Remove();

        // Set the display type of the top‑level equation (preserve formatting).
        officeMath.DisplayType = OfficeMathDisplayType.Display;

        // Second equation ------------------------------------------------------
        builder.Writeln();
        builder.Writeln("Second equation:");

        Field field2 = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq2 = field2 as FieldEQ;
        if (fieldEq2 == null)
            throw new InvalidOperationException("Failed to create second FieldEQ.");

        builder.MoveTo(fieldEq2.Separator);
        builder.Write(@"\r(3,x)"); // Simple square root expression

        // Update the second field before conversion.
        field2.Update();

        OfficeMath officeMath2 = fieldEq2.AsOfficeMath();
        if (officeMath2 == null)
            throw new InvalidOperationException("Second EQ field could not be converted to OfficeMath.");

        Node field2Start = fieldEq2.Start;
        field2Start.ParentNode.InsertBefore(officeMath2, field2Start);
        fieldEq2.Remove();

        // Ensure the second equation is also displayed as a block.
        officeMath2.DisplayType = OfficeMathDisplayType.Display;

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document as DOCX.
        string outputPath = Path.Combine(outputDir, "ModifiedDocument.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Reload the saved document to verify equations are preserved.
        Document loadedDoc = new Document(outputPath);
        NodeCollection mathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes == null || mathNodes.Count < 2)
            throw new InvalidOperationException("Expected OfficeMath equations were not found after saving.");

        // Verify that each top‑level equation retains its display type.
        foreach (OfficeMath om in mathNodes)
        {
            // Only top‑level equations have MathObjectType == OMathPara.
            if (om.MathObjectType == MathObjectType.OMathPara && om.DisplayType != OfficeMathDisplayType.Display)
                throw new InvalidOperationException("Equation display type was not preserved.");
        }

        Console.WriteLine("Document saved successfully with preserved OfficeMath equations.");
    }
}
