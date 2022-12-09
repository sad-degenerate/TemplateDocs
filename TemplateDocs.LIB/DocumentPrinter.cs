using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;
using Point = System.Drawing.Point;
using Task = System.Threading.Tasks.Task;

namespace TemplateDocs.LIB
{
    public class DocumentPrinter
    {
        /// <summary>
        /// Путь к файлу для печати.
        /// </summary>
        private readonly string _printFilePath;

        /// <summary>
        /// Создать новый экземпляр класса DocumentPrinter.
        /// </summary>
        /// <param name="resultFilePath"></param>
        /// <exception cref="ArgumentNullException">Если путь к файлу пуст.</exception>
        /// <exception cref="FileNotFoundException">Если по указанному пути файла не существует.</exception>
        /// <exception cref="ArgumentException">Если файл имеет недопустимое расширение (допустимое - только ".docx").</exception>
        public DocumentPrinter(string resultFilePath)
        {
            if (string.IsNullOrWhiteSpace(resultFilePath))
                throw new ArgumentNullException(nameof(resultFilePath), "Путь к файлу для печати пуст.");
            if (File.Exists(resultFilePath) == false)
                throw new FileNotFoundException("Указанного вами пути не существует.", nameof(resultFilePath));
            if (Path.GetExtension(resultFilePath) != ".docx")
                throw new ArgumentException("Файл должен иметь расширение \"docx\".", nameof(resultFilePath));

            _printFilePath = resultFilePath;
        }

        /// <summary>
        /// Метод, печатающий документ.
        /// </summary>
        /// <param name="copies">Количество копий документа.</param>
        /// <param name="isDuplex">Использовать ли двухстороннюю печать.</param>
        public async Task PrintAsync(int copies, bool isDuplex)
        {
            await Task.Run(() => Print(copies, isDuplex));
        }

        /// <summary>
        /// Метод, печатающий документ.
        /// </summary>
        /// <param name="copies">Количество копий документа.</param>
        /// <param name="isDuplex">Использовать ли двухстороннюю печать.</param>
        public void Print(int copies, bool isDuplex)
        {
            var images = GenerateImages();

            var print = new PrintDocument();
            var currentImage = 0;

            print.PrintPage += (o, e) =>
            {
                e.Graphics.DrawImage(images[currentImage], new Point(0, 0));
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

            print.PrinterSettings.Duplex = isDuplex ? Duplex.Vertical : Duplex.Simplex;

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

            foreach (var bits in from Window window in doc.Windows 
                     from Pane pane in window.Panes 
                     from Page page in pane.Pages select page.EnhMetaFileBits)
            {
                try
                {
                    using (var ms = new MemoryStream((byte[])(bits)))
                    {
                        printList.Add(Image.FromStream(ms));
                    }
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Не удалось прочитать документ для печати.");
                }
            }

            app.Quit();

            return printList;
        }
    }
}