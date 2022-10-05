using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;
// TODO: 01: Bring the Microsoft.Office.Interop.Word namespace into scope.



namespace FourthCoffee.ExceptionLogger
{
    /// <summary>
    /// Represents the <see cref="FourthCoffee.ExceptionLogger.ExceptionCombiner" /> class in the object model.
    /// </summary>
    public class ExceptionCombiner
    {
        string _outputFilePath;
        IEnumerable<ExceptionEntry> _exceptions;
        private Application _word;

        // TODO: 02: Declare a global object to encapsulate Microsoft Word.


        public ExceptionCombiner(string outputFilePath, IEnumerable<ExceptionEntry> exceptions)
        {
            if (exceptions == null)
                throw new NullReferenceException("exceptions");

            _outputFilePath = outputFilePath;
            _exceptions = exceptions;
        }

        public void WriteToWordDocument()
        {
            if (_exceptions.Count() == 0)
                return;

            // TODO: 03: Instantiate the _word object.
            _word = new Application();
            

            GenerateWordDocument();

            _word.Quit();
      
        }

        private void GenerateWordDocument()
        {
            CreateDocument();

            AppendText("Exception Report", true, true);

            InsertCarriageReturn();

            var count = 1;
            foreach (var exception in _exceptions)
            {
                AppendText(
                    string.Format("{0}) {1}", count, exception.Title), 
                    true, 
                    false);

                InsertCarriageReturn();
                AppendText(exception.Details, false, false);
                InsertCarriageReturn();
                InsertCarriageReturn();
                count++;
            }

            Save();          
        }

        private void CreateDocument()
        {
            // TODO: 04: Create a blank Word document.
            _word.Documents.Add().Activate();
        }

        private void AppendText(string text, bool bold, bool underLine)
        {
            var currentLocation = GetEndOfDocument();

            currentLocation.Bold = bold ? 1 : 0;
            currentLocation.Underline = underLine ?
               WdUnderline.wdUnderlineSingle :
               WdUnderline.wdUnderlineNone;
            currentLocation.InsertAfter(text);      
        }

        private void InsertCarriageReturn()
        {
            var currentLocation = GetEndOfDocument();

            currentLocation.InsertBreak(WdBreakType.wdLineBreak);
       
        }

        private void Save()
        {
            if (File.Exists(_outputFilePath))
                File.Delete(_outputFilePath);

            var currentDocument = _word.ActiveDocument;

            currentDocument.SaveAs(_outputFilePath);
            currentDocument.Close();
     
        }

        private Range GetEndOfDocument()
        {
            var endOfDocument = _word.ActiveDocument.Content.End - 1;
            return _word.ActiveDocument.Range(endOfDocument); 
        }
    }
}
