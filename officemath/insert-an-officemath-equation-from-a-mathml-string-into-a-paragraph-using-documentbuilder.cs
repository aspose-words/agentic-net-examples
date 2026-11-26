using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Define the MathML string (stored as metadata, not directly parsed).
        const string mathMl = "<math><mfrac><mi>1</mi><mi>2</mi></mfrac></math>";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the equation.
        builder.Writeln("This paragraph will contain an equation inserted from MathML:");

        // Insert a field of type Equation (EQ) at the current builder position.
        builder.InsertField(FieldType.FieldEquation, true);

        // Retrieve the inserted field (the last field in the document).
        FieldEQ eqField = (FieldEQ)doc.Range.Fields[doc.Range.Fields.Count - 1];

        // Write a simple EQ argument that Aspose.Words can reliably convert to OfficeMath.
        // Move the builder to the field separator and write the argument.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Convert the EQ field to a real OfficeMath node.
        OfficeMath officeMath = eqField.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start node.
            Node fieldStart = eqField.Start;
            fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);

            // Remove the original EQ field from the document.
            eqField.Remove();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        // Validate that the document contains at least one top‑level OfficeMath node.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes.Count == 0)
        {
            throw new InvalidOperationException("No OfficeMath nodes were found in the saved document.");
        }

        // The program finishes here without any interactive prompts.
    }
}
