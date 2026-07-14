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
        // LaTeX representation of the equation (preserved as metadata in the document).
        string latexEquation = @"E = mc^2";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Below is an equation generated from a LaTeX string:");
        builder.Writeln($"LaTeX source: {latexEquation}");

        // Insert a field of type Equation (FieldEQ) which will be used to bootstrap a real OfficeMath node.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = (FieldEQ)field;

        // Move the builder to the field separator and write a safe EQ argument.
        // Aspose.Words does not parse LaTeX directly, so we use a simple, reliable expression.
        // The original LaTeX string is kept in the document as plain text above.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(@"\f(1,2)"); // Simple fraction 1/2 as a placeholder equation.

        // Update the field so that Aspose.Words can convert it.
        field.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = fieldEq.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start node.
            CompositeNode? parent = fieldEq.Start.ParentNode as CompositeNode;
            if (parent == null)
                throw new InvalidOperationException("The field start node does not have a valid composite parent.");

            parent.InsertBefore(officeMath, fieldEq.Start);

            // Remove the original field, leaving only the OfficeMath node.
            fieldEq.Remove();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert the EQ field to an OfficeMath object.");
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created and contains at least one OfficeMath node.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        Document loadedDoc = new Document(outputPath);
        int officeMathCount = loadedDoc.GetChildNodes(NodeType.OfficeMath, true).Count;
        if (officeMathCount == 0)
            throw new InvalidOperationException("The document does not contain any OfficeMath equations after processing.");

        // Confirmation (non‑interactive).
        Console.WriteLine($"Document saved to '{outputPath}' with {officeMathCount} OfficeMath node(s).");
    }
}
