using System.IO;

namespace TemplateDocs.LIB
{
    public class Document
    {
        public FileInfo DocumentInfo { get; private set; }

        public Document(string path)
        {
            if (File.Exists(path) == false)
                throw new FileNotFoundException("Не удалось открыть файл.");

            DocumentInfo = new FileInfo(path);
        }
    }
}