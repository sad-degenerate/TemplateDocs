using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;

namespace TemplateDocs.LIB
{
    public class DocumentPrinter
    {
        /// <summary>
        /// Путь к файлу для печати.
        /// </summary>
        private string _printFilePath;

        /// <summary>
        /// Создать новый экземпляр класса DocumentPrinter.
        /// </summary>
        /// <param name="resultFilePath"></param>
        /// <exception cref="ArgumentNullException">Если путь к файлу пуст.</exception>
        /// <exception cref="FileNotFoundException">Если по указанному пути файла не существует.</exception>
        /// <exception cref="ArgumentException">Если файл имеет недопустимое расширение (допустимое - только ".docx").</exception>
        public DocumentPrinter(string resultFilePath)
        {
            if (string.IsNullOrWhiteSpace(resultFilePath) == true)
                throw new ArgumentNullException(nameof(resultFilePath), "Путь к файлу для печати пуст.");
            if (File.Exists(resultFilePath) == false)
                throw new FileNotFoundException("Указанного вами пути не существует.", nameof(resultFilePath));
            if (Path.GetExtension(resultFilePath) != ".docx")
                throw new ArgumentException("Файл должен иметь расширение \"docx\".", nameof(resultFilePath));
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
            var doc = app.Documents.Open(_printFilePath);
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