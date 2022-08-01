using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;

namespace TemplateDocs.LIB
{
    public class DocumentManager
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
        /// Путь к документу, собранному по шаблону.
        /// </summary>
        private string _resultFile;

        /// <summary>
        /// Создать новый объект класса Document.
        /// </summary>
        /// <param name="path">Путь к документу, в котором находится шаблон.</param>
        /// <param name="outputPath">Путь к папке, в которой будут находиться готовые документы.</param>
        public DocumentManager(string path, string outputPath)
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

            _resultFile = Path.Combine(_outputPath, documentName);
            File.Copy(_templateDoc.FullName, _resultFile, true);

            ReplaceWords(replaceWords, _resultFile);
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

        /// <summary>
        /// Метод, печатающий документ с результатами программы.
        /// </summary>
        /// <param name="copies">Количество копий документа.</param>
        public void Print(int copies)
        {
            var images = GenerateImages();

            var print = new PrintDocument();
            int currentImage = 0;

            print.PrintPage += (o, e) =>
            {
                e.Graphics.DrawImage(images[currentImage], new System.Drawing.Point(0, 0));
                currentImage++;
                if (images.Count <= currentImage)
                {
                    if (copies > 0)
                    {
                        copies--;
                        currentImage = 0;
                        e.HasMorePages = true;
                    }
                        
                    e.HasMorePages = false;
                }
                else
                    e.HasMorePages = true;
            };

            print.PrinterSettings.Duplex = Duplex.Vertical;

            print.Print();
        }

        /// <summary>
        /// Метод, который переводит документ Word в серию изображений.
        /// </summary>
        /// <returns>Список изображений, полученных из страниц документа.</returns>
        /// <exception cref="ArgumentException">Ошибка, возникающая, если не удалось преобразовать документ в список изображений.</exception>
        private List<Image> GenerateImages()
        {
            var printList = new List<Image>();
            var app = new Application();
            var doc = app.Documents.Open(_resultFile);
            app.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
            app.Visible = false;

            foreach (Window window in doc.Windows)
            {
                foreach (Pane pane in window.Panes)
                {
                    foreach (Page page in pane.Pages)
                    {
                        var bits = page.EnhMetaFileBits;

                        try
                        {
                            using (var ms = new MemoryStream((byte[])(bits)))
                                printList.Add(Image.FromStream(ms));
                        }
                        catch (Exception ex)
                        {
                            throw new ArgumentException("Не удалось прочитать документ для печати.");
                        }
                    }
                }
            }

            app.Quit();

            return printList;
        }
    }
}