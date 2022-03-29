using System;
using System.Collections.Generic;
using System.IO;

namespace TemplateDocs.LIB
{
    public class ReplaceWords
    {
        public Dictionary<string, string> Words { get; }

        public ReplaceWords(Dictionary<string, string> words)
        {
            if (words == null || words.Count == 0)
                throw new ArgumentException("Передан пустой или неопределенный список.", nameof(words));

            Words = words;
        }

        public ReplaceWords(FileInfo file)
        {
            if (file == null)
                throw new ArgumentNullException(nameof(file), "Не передан файл для считывания слов для замены.");
            if (file.Exists == false || Path.GetExtension(file.FullName) != ".txt")
                throw new ArgumentException("По указанному пути нет подходящего файла.", nameof(file));

            using (StreamReader sr = new StreamReader(file.FullName))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                    GetWordsFromLine(line);
            }
        }

        private void GetWordsFromLine(string line, string separator = "|")
        {
            var words = line.Split('|');

            if (words.Length == 2)
                Words.Add(words[0], words[1]);
            else
                throw new ArgumentException("Не удалось считать данные, проверьте правильность заполнения файла.", nameof (words));
        }
    }
}