using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using TemplateDocs.LIB;

namespace TemplateDocs.Test
{
    [TestClass]
    public class ReplaceWordsTest
    {
        [TestMethod]
        public void ReplaceWordsBase()
        {
            var filePath = $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\replaceWords.txt";
            var file = File.Create(filePath);
            file.Close();
            var fileInfo = new FileInfo(filePath);
            var fakeFileInfo = new FileInfo($"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\replaceWords1.txt");

            // Файла не существует

            try
            {
                var replaceWords = new ReplaceWords(fakeFileInfo);
                Assert.Fail();
            }
            catch (Exception ex)
            {

            }

            // Файл пуст

            try
            {
                var replaceWords = new ReplaceWords(fileInfo);
                Assert.Fail();
            }
            catch (Exception ex)
            {

            }

            // В файле написано не то что нужно

            using (var sw = new StreamWriter(fileInfo.FullName))
            {
                sw.Write("РАБОТА");
            }

            try
            {
                var replaceWords = new ReplaceWords(fileInfo);
                Assert.Fail();
            }
            catch (Exception ex)
            {

            }

            // Правильная работа программы

            using (var sw = new StreamWriter(fileInfo.FullName))
            {
                sw.WriteLine("|ЗАБОТА");
                sw.WriteLine("АПТЕКА|КОТ");
            }

            var replaceWordsFin = new ReplaceWords(fileInfo);
            Assert.IsTrue(replaceWordsFin.Words.Count == 2);
        }
    }
}