using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using TemplateDocs.LIB;

namespace TemplateDocs.Test
{
    [TestClass]
    public class DocumentTest
    {
        [TestMethod]
        public void DocumentTestMethod()
        {
            // Файл существует.

            var path = $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\newFile.docx";

            File.Create(path);

            try
            {
                var trueDoc = new Document(path);
            }
            catch (Exception ex)
            {
                Assert.Fail();
            }

            // Файл не существует.

            path = $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\newFile1.docx";

            try
            {
                var fakeDoc = new Document(path);

                Assert.Fail();
            }
            catch (Exception ex)
            {
                
            }

            // У файла другое расширение.

            path = $"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\\newFile.txt";

            File.Create(path);

            try
            {
                var fakeDoc = new Document(path);

                Assert.Fail();
            }
            catch (Exception ex)
            {

            }
        }
    }
}