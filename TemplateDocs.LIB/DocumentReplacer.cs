using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;

namespace TemplateDocs.LIB
{
    public class DocumentReplacer
    {
        /// <summary>
        /// Путь к шаблону для создания документов.
        /// </summary>
        private FileInfo _templateDoc;
        /// <summary>
        /// Путь к документу, собирающемуся по шаблону.
        /// </summary>
        private string _outputPath;

        /// <summary>
        /// Создать новый объект класса Document.
        /// </summary>
        /// <param name="path">Путь к документу, в котором находится шаблон.</param>
        /// <param name="outputPath">Путь к папке, в которой будут находиться готовые документы.</param>
        public DocumentReplacer(string path, string outputPath)
        {
            if (Directory.Exists(outputPath) == false)
                Directory.CreateDirectory(outputPath);
            if (File.Exists(path) == false)
                throw new FileNotFoundException("Не удалось открыть файл.");
            if (Path.GetExtension(path) != ".docx")
                throw new ArgumentException("Файл должен иметь расширение \"docx\".", nameof(path));
            
            _outputPath = outputPath;
            _templateDoc = new FileInfo(path);
        }

        /// <summary>
        /// Точка запуска программы, создает новый файл по указанному пути,
        /// в котором произведена замена по шаблону.
        /// </summary>
        /// <param name="replaceWords">Список из слов для замены, в котором 
        /// Key: слово, подлежащее замене, Value: слово, которое встанет на его место.</param>
        /// <param name="documentName">Название нового файла, в котором будет произведена замена.</param>
        public void Replace(Dictionary<string, string> replaceWords, string documentName)
        {
            if (documentName.EndsWith(".docx") == false)
                documentName += ".docx";

            var destFile = Path.Combine(_outputPath, documentName);
            File.Copy(_templateDoc.FullName, destFile, true);

            ReplaceWords(replaceWords, destFile);
        }

        /// <summary>
        /// Метод производящий замену слов.
        /// </summary>
        /// <param name="replaceWords">Список слов, для замены.</param>
        /// <param name="filePath">Путь к файлу, в которому нужно заменить слова.</param>
        private void ReplaceWords(Dictionary<string, string> replaceWords, string filePath)
        {
            var app = new Application();

            app.Documents.Open(filePath);

            foreach (var word in replaceWords)
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

            app.Documents.Save();
            app.Quit();
        }

        public void Print(int copies = 1)
        {
            var app = new Application();
            var document = app.Documents.Open(_templateDoc.FullName);

            document.PrintOut(true, false, WdPrintOutRange.wdPrintAllDocument, Item: WdPrintOutItem.wdPrintDocumentContent,
                                Copies: "1", Pages: "", PageType: WdPrintOutPages.wdPrintAllPages, PrintToFile: false,
                                Collate: true, ManualDuplexPrint: false);
        }
    }
}