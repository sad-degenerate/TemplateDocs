using System;
using System.Collections.Generic;

namespace TemplateDocs.LIB
{
    public class ReplaceWords
    {
        public Dictionary<string, string> Words { get; private set; }

        public ReplaceWords(Dictionary<string, string> words)
        {
            if (words.Count == 0)
                throw new ArgumentNullException(nameof(words), "Пустой список слов на замену.");

            Words = words;
        }
    }
}