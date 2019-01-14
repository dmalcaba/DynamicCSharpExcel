using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DynamicCSharpWord
{

    public class Word
    {
        private dynamic _wordObj;

        /// <summary>
        /// Version information, use for logging purposes in case of errors
        /// </summary>
        public string Version => $"{_wordObj.Name} Version {_wordObj.Version} Build {_wordObj.Build}";

        /// <summary>
        /// Word Mail Merge
        /// https://www.youtube.com/watch?v=yj6mIe8cyZ8
        /// </summary>

        public Word()
        {
            Type wordType = Type.GetTypeFromProgID("Word.Application", true);
            _wordObj = Activator.CreateInstance(wordType);

            dynamic documentObj = _wordObj.Documents.Add();

            _wordObj.Visible = true;

        }
    }
}
