using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder for it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper method that inserts a simple equation using the deterministic EQ‑field bootstrap workflow.
        void InsertEquation(string eqArgument)
        {
            // Insert an Equation field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);

            // Move to the field separator and write the EQ argument.
            builder.MoveTo(field.Separator);
            builder.Write(eqArgument);

            // Update the field so that Aspose.Words can convert it to a real OfficeMath node.
            field.Update();

            // Convert the field to an OfficeMath node.
            FieldEQ fieldEq = field as FieldEQ;
            if (fieldEq == null)
                throw new InvalidOperationException("Failed to cast field to FieldEQ.");

            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath == null)
                throw new InvalidOperationException("EQ field conversion returned null OfficeMath.");

            // Insert the OfficeMath node before the field start and remove the original field.
            CompositeNode parent = field.Start.ParentNode as CompositeNode;
            if (parent == null)
                throw new InvalidOperationException("Field start does not have a composite parent.");

            parent.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Insert first equation.
        builder.Writeln("First equation:");
        InsertEquation(@"\f(1,2)"); // Simple fraction.

        // Insert second equation.
        builder.Writeln("Second equation:");
        InsertEquation(@"\f(3,4)"); // Another simple fraction.

        // Save the sample document.
        const string samplePath = "Sample.docx";
        doc.Save(samplePath, SaveFormat.Docx);

        // Reload the document to simulate a typical load‑modify‑save scenario.
        Document loadedDoc = new Document(samplePath);

        // Set display mode of all top‑level OfficeMath equations to Inline.
        var topLevelEquations = loadedDoc.GetChildNodes(NodeType.OfficeMath, true)
                                         .Cast<OfficeMath>()
                                         .Where(om => om.MathObjectType == MathObjectType.OMathPara);

        foreach (OfficeMath om in topLevelEquations)
        {
            om.DisplayType = OfficeMathDisplayType.Inline;
        }

        // Save the modified document.
        const string outputPath = "Output.docx";
        loadedDoc.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file exists.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Reload the output document and verify that all top‑level equations are Inline.
        Document resultDoc = new Document(outputPath);
        var resultEquations = resultDoc.GetChildNodes(NodeType.OfficeMath, true)
                                      .Cast<OfficeMath>()
                                      .Where(om => om.MathObjectType == MathObjectType.OMathPara);

        foreach (OfficeMath om in resultEquations)
        {
            if (om.DisplayType != OfficeMathDisplayType.Inline)
                throw new InvalidOperationException("An equation's DisplayType is not set to Inline.");
        }

        // If we reach this point, the operation succeeded.
        Console.WriteLine("All top‑level OfficeMath equations set to Inline and saved successfully.");
    }
}
