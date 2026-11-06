using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the template and output documents.
        string templatePath = "Template.docx";
        string outputPath = "Output.docx";

        // Create a new document and add a few equations using the deterministic EQ‑field bootstrap workflow.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert first paragraph with an equation.
        builder.Writeln("First equation:");
        InsertEquation(builder, @"\f(1,2)");

        // Insert second paragraph with an equation.
        builder.Writeln();
        builder.Writeln("Second equation:");
        InsertEquation(builder, @"\r(3,x)");

        // Insert third paragraph with an equation.
        builder.Writeln();
        builder.Writeln("Third equation:");
        InsertEquation(builder, @"\f(a,b)");

        // Save the template document.
        doc.Save(templatePath, SaveFormat.Docx);

        // Load the document (simulating a template scenario).
        Document loadedDoc = new Document(templatePath);

        // Update justification of all top‑level OfficeMath equations to right alignment.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in officeMathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                om.Justification = OfficeMathJustification.Right;
            }
        }

        // Save the modified document.
        loadedDoc.Save(outputPath, SaveFormat.Docx);

        // Validation: ensure the output file exists.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The output file '{outputPath}' was not created.");
        }

        // Validation: ensure all top‑level OfficeMath nodes have right justification.
        Document verifyDoc = new Document(outputPath);
        NodeCollection verifyNodes = verifyDoc.GetChildNodes(NodeType.OfficeMath, true);
        foreach (OfficeMath om in verifyNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara &&
                om.Justification != OfficeMathJustification.Right)
            {
                throw new InvalidOperationException("One or more OfficeMath equations do not have right justification.");
            }
        }

        // Indicate successful completion (no interactive output required).
        Console.WriteLine("Justification update completed successfully.");
    }

    // Helper method to insert an equation using the EQ‑field bootstrap workflow.
    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        if (field is FieldEQ fieldEq)
        {
            // Move the builder to the field separator and write the EQ argument.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(eqArgument);

            // Update the field so that Aspose can convert it to OfficeMath.
            fieldEq.Update();

            // Convert the field to an OfficeMath object.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                Node fieldStart = field.Start;
                Node parent = fieldStart.ParentNode;
                if (parent is CompositeNode compositeParent)
                {
                    compositeParent.InsertBefore(officeMath, fieldStart);
                }
                else
                {
                    throw new InvalidOperationException("Unexpected parent node type for the field start.");
                }

                // Remove the original field.
                field.Remove();
            }
            else
            {
                throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
            }
        }
        else
        {
            throw new InvalidOperationException("Inserted field is not a FieldEQ.");
        }
    }
}
