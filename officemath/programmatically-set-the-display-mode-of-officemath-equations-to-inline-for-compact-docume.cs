using System;
using System.IO;
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

        // Insert a paragraph with introductory text.
        builder.Writeln("Sample document with inline equations:");

        // ---------- Bootstrap first equation ----------
        // Insert an EQ field.
        FieldEQ field1 = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        // Move to the field separator and write a simple fraction equation.
        builder.MoveTo(field1.Separator);
        builder.Write(@"\f(1,2)"); // 1/2
        // Return the builder to the paragraph after the field.
        builder.MoveTo(field1.Start.ParentNode);
        builder.InsertParagraph();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath1 = field1.AsOfficeMath();
        if (officeMath1 != null)
        {
            // Insert the OfficeMath node before the field start.
            field1.Start.ParentNode.InsertBefore(officeMath1, field1.Start);
            // Remove the original field.
            field1.Remove();
        }

        // ---------- Bootstrap second equation ----------
        FieldEQ field2 = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
        builder.MoveTo(field2.Separator);
        builder.Write(@"\r(3,x)"); // Cube root of x
        builder.MoveTo(field2.Start.ParentNode);
        builder.InsertParagraph();

        OfficeMath officeMath2 = field2.AsOfficeMath();
        if (officeMath2 != null)
        {
            field2.Start.ParentNode.InsertBefore(officeMath2, field2.Start);
            field2.Remove();
        }

        // ---------- Set display mode to Inline for all top‑level equations ----------
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in mathNodes)
        {
            // Only modify top‑level OfficeMath (MathObjectType.OMathPara).
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.DisplayType = OfficeMathDisplayType.Inline;
                // Justification must be Inline when DisplayType is Inline.
                om.Justification = OfficeMathJustification.Inline;
            }
        }

        // Save the document.
        string outputPath = "InlineEquations.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // The program finishes automatically.
    }
}
