using System;
using Word = Microsoft.Office.Interop.Word;

class Program
{
    /// <summary>
    /// Launches Microsoft Word and types a sample letter to the editor
    /// character by character to mimic user input.
    /// </summary>
    static void Main(string[] args)
    {
        // Start Word and make it visible
        var wordApp = new Word.Application();
        wordApp.Visible = true;

        // Create a new document
        var document = wordApp.Documents.Add();

        string[] lines = new[]
        {
            "Dear Editor,",
            "",
            "I am writing to express my views on the recent events.",
            "Thank you for considering my letter.",
            "",
            "Sincerely,",
            "A Concerned Reader"
        };

        foreach (var line in lines)
        {
            foreach (char c in line)
            {
                wordApp.Selection.TypeText(c.ToString());
                System.Threading.Thread.Sleep(50); // small delay to mimic typing
            }
            wordApp.Selection.TypeParagraph();
        }
    }
}
