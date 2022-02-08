using Microsoft.Office.Interop.Word;
using System;

namespace TemplateDocs.LIB
{
    public static class ReplaceInDoc
    {
        public static void Replace(ReplaceWords words, Document doc)
        {
            var app = new Application();
            object file = doc.DocumentInfo.FullName;

            app.Documents.Open(file);

            foreach (var word in words.Words)
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
        }
    }
}