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

        // Add introductory text.
        builder.Writeln("Paragraph before the equation.");

        // -----------------------------------------------------------------
        // Insert an OfficeMath equation that originates from a MathML string.
        // The MathML is kept as a comment for reference only.
        // -----------------------------------------------------------------
        // MathML source (kept for documentation purposes only):
        // <math><mi>x</mi><mo>=</mo><mfrac><mi>-b</mi><mi>2a</mi></mfrac></math>
        string mathML = "<math><mi>x</mi><mo>=</mo><mfrac><mi>-b</mi><mi>2a</mi></mfrac></math>";

        // Insert a deterministic EQ field that will be converted to OfficeMath.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = field as FieldEQ;
        if (fieldEQ == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write a simple EQ argument that Aspose.Words can reliably convert.
        // Using a fraction: \f(1,2)
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that its result reflects the written argument.
        field.Update();

        // Convert the field to a real OfficeMath node.
        OfficeMath officeMath = fieldEQ.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field did not convert to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the field.
        Node fieldStart = fieldEQ.Start;
        Node parentNode = fieldStart.ParentNode;
        if (parentNode is CompositeNode parentComposite)
        {
            parentComposite.InsertBefore(officeMath, fieldStart);
        }
        else
        {
            throw new InvalidOperationException("Field start does not have a composite parent.");
        }

        // Remove the original EQ field from the document.
        fieldEQ.Remove();

        // Continue with normal document content.
        builder.Writeln();
        builder.Writeln("Paragraph after the equation.");

        // Save the document.
        string outputPath = "OfficeMathFromMathML.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created and contains an OfficeMath node.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Output file was not created.");

        Document loaded = new Document(outputPath);
        NodeCollection officeMathNodes = loaded.GetChildNodes(NodeType.OfficeMath, true);
        if (officeMathNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath nodes found in the saved document.");
    }
}
