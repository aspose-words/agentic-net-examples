using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Loading;

class OfficeMathTypeCollector : DocumentVisitor
{
    // Stores the MathObjectType of each visited OfficeMath node.
    public readonly List<MathObjectType> Types = new List<MathObjectType>();

    // Called when an OfficeMath node is encountered.
    public override VisitorAction VisitOfficeMathStart(OfficeMath officeMath)
    {
        Types.Add(officeMath.MathObjectType);
        return VisitorAction.Continue;
    }
}

class Program
{
    static void Main()
    {
        // Path to the Markdown document.
        const string inputPath = "input.md";

        // Load the document. LoadOptions can be customized if needed.
        LoadOptions loadOptions = new LoadOptions();
        Document doc = new Document(inputPath, loadOptions);

        // Traverse the document and collect MathObjectTypes.
        OfficeMathTypeCollector collector = new OfficeMathTypeCollector();
        doc.Accept(collector);

        // Output the collected MathObjectTypes.
        foreach (MathObjectType type in collector.Types)
        {
            Console.WriteLine(type);
        }
    }
}
