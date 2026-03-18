using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Bibliography;

class RetrieveFourthPerson
{
    static void Main()
    {
        // Load an existing Word document that contains bibliography sources.
        Document doc = new Document("BibliographySources.docx");

        // Access the bibliography of the document.
        Bibliography bibliography = doc.Bibliography;

        // Get the first source in the bibliography (adjust as needed).
        Source source = bibliography.Sources.FirstOrDefault();
        if (source == null)
        {
            Console.WriteLine("No bibliography sources found.");
            return;
        }

        // The 'Author' contributor holds a PersonCollection.
        PersonCollection authors = (PersonCollection)source.Contributors.Author;

        // Ensure the collection has at least four persons.
        if (authors.Count < 4)
        {
            Console.WriteLine("The bibliography source does not contain four authors.");
            return;
        }

        // Retrieve the fourth person (zero‑based index 3) using LINQ's ElementAt.
        Person fourthPerson = authors.ElementAt(3);

        // Display the person's full name.
        Console.WriteLine($"Fourth author: {fourthPerson.First} {fourthPerson.Middle} {fourthPerson.Last}");

        // Optionally, save the document after any modifications.
        doc.Save("BibliographySources_Processed.docx");
    }
}
