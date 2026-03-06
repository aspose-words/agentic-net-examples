using System;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Access the mail merge settings of the document.
        MailMergeSettings settings = doc.MailMergeSettings;

        // Evaluate and display the current mail merge main document type.
        Console.WriteLine("Current MainDocumentType: " + settings.MainDocumentType);

        // Evaluate and display the current mail merge destination.
        Console.WriteLine("Current Destination: " + settings.Destination);

        // Display additional mail merge settings for completeness.
        Console.WriteLine("DataSource: " + settings.DataSource);
        Console.WriteLine("DataType: " + settings.DataType);
        Console.WriteLine("CheckErrors: " + settings.CheckErrors);
        Console.WriteLine("DoNotSupressBlankLines: " + settings.DoNotSupressBlankLines);
        Console.WriteLine("ViewMergedData: " + settings.ViewMergedData);

        // List all possible values of MailMergeMainDocumentType.
        Console.WriteLine("\nAll MailMergeMainDocumentType values:");
        foreach (MailMergeMainDocumentType type in Enum.GetValues(typeof(MailMergeMainDocumentType)))
        {
            Console.WriteLine($"{type} = {(int)type}");
        }

        // List all possible values of MailMergeDestination.
        Console.WriteLine("\nAll MailMergeDestination values:");
        foreach (MailMergeDestination dest in Enum.GetValues(typeof(MailMergeDestination)))
        {
            Console.WriteLine($"{dest} = {(int)dest}");
        }

        // Example modification: set the document to be a form letter and output to a new document.
        settings.MainDocumentType = MailMergeMainDocumentType.FormLetters;
        settings.Destination = MailMergeDestination.NewDocument;

        // Save the modified document to a new file.
        doc.Save("Evaluated.docx");
    }
}
