using Microsoft.Office.Interop.Word;
using System;
using System.IO;

namespace TemplateDocs.LIB
{
    public class Document
    {
        private FileInfo _documentInfo;

        public Document(string path)
        {
            SetActiveDocument(path);
        }

        private void SetActiveDocument(string path)
        {
            if (File.Exists(path) == false)
                throw new FileNotFoundException("Не удалось открыть файл.");
            if (Path.GetExtension(path) != ".docx")
                throw new ArgumentException("Файл должен иметь расширение \"docx\".", nameof(path));

            _documentInfo = new FileInfo(path);
        }

        public void CopyDocumentAndActivate(string name, string outputPath)
        {
            var destFile = outputPath + name;
            File.Copy(_documentInfo.FullName, destFile);
            SetActiveDocument(destFile);
        }

        public void ReplaceWords(ReplaceWords words)
        {
            var app = new Application();
            object file = _documentInfo.FullName;

            app.Documents.Open(file);

            foreach (var word in words.Words)
            {
                app.Selection.Find.Execute(FindText: word.Key,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: Type.Missing,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: WdFindWrap.wdFindContinue,
                    Format: false,
                    ReplaceWith: word.Value,
                    Replace: WdReplace.wdReplaceAll);
            }

            app.Documents.Close(file);
        }
    }
}