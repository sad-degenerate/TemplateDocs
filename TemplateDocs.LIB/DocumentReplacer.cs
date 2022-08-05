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
        public string OutputDocumentPath { get; private set; }

        /// <summary>
        /// Создать новый объект класса Document.
        /// </summary>
        /// <param name="path">Путь к документу, в котором находится шаблон.</param>
        /// <param name="outputPath">Путь к папке, в которой будут находиться готовые документы.</param>
        /// <exception cref="FileNotFoundException">Если по указанному пути не найден файл.</exception>
        /// <exception cref="ArgumentException">Если расширение файла не соответствует ".docx".</exception>
        public DocumentReplacer(string path, string outputPath)
        {
            if (Directory.Exists(outputPath) == false)
                Directory.CreateDirectory(outputPath);
            if (File.Exists(path) == false)
                throw new FileNotFoundException("Не удалось открыть файл.");
            if (Path.GetExtension(path) != ".docx")
                throw new ArgumentException("Файл должен иметь расширение \"docx\".", nameof(path));
            
            OutputDocumentPath = outputPath;
            _templateDoc = new FileInfo(path);
        }

        /// <summary>
        /// Точка запуска программы, создает новый файл по указанному пути,
        /// в котором произведена замена по шаблону.
        /// </summary>
        /// <param name="replaceWords">Список из слов для замены, в котором 
        /// Key: слово, подлежащее замене, Value: слово, которое встанет на его место.</param>
        /// <param name="documentName">Название нового файла, в котором будет произведена замена.</param>
        /// <returns>Путь к файлу, в котором произошла замена.</returns>
        public string Replace(Dictionary<string, string> replaceWords, string documentName)
        {
            if (Path.GetExtension(documentName) != ".docx")
                documentName += ".docx";

            var resultFilePath = Path.Combine(OutputDocumentPath, documentName);
            File.Copy(_templateDoc.FullName, resultFilePath, true);

            ReplaceWords(replaceWords, resultFilePath);
            return resultFilePath;
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
    }
}