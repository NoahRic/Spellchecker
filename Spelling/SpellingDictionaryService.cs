//***************************************************************************
//
//    Copyright (c) Microsoft Corporation. All rights reserved.
//    This code is licensed under the Visual Studio SDK license terms.
//    THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
//    ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
//    IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
//    PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
//
//***************************************************************************

﻿////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright (c) Microsoft Corporation.  All rights reserved.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Windows.Controls;
using Microsoft.VisualStudio.Text;

namespace Microsoft.VisualStudio.Language.Spellchecker
{
    /// <summary>
    /// Spell checking provider based on WPF spell checker
    /// </summary>
    [Export(typeof(ISpellingDictionaryService))]
    internal class SpellingDictionaryService : ISpellingDictionaryService
    {
        #region Private data
        private SortedSet<string> _ignoreWords = new SortedSet<string>();
        private string _ignoreWordsFile;
        #endregion

        #region Constructor
        /// <summary>
        /// Constructor for SpellingDictionaryService
        /// </summary>
        public SpellingDictionaryService()
        {
            string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\VisualStudio\10.0\SpellChecker");
            if (!Directory.Exists(localFolder))
            {
                Directory.CreateDirectory(localFolder);
            }
            _ignoreWordsFile = Path.Combine(localFolder, "Dictionary.txt");

            LoadIgnoreDictionary();
        }
        #endregion

        #region ISpellingDictionaryService

        /// <summary>
        /// Adds given word to the dictionary.
        /// </summary>
        /// <param name="word">The word to add to the dictionary.</param>
        public void AddWordToDictionary(string word)
        {
            if (!string.IsNullOrEmpty(word))
            {
                // Add this word to the dictionary file.
                using (StreamWriter writer = new StreamWriter(_ignoreWordsFile, true))
                {
                    writer.WriteLine(word);
                }

                IgnoreWord(word);
            }
        }

        public void IgnoreWord(string word)
        {
            if (!string.IsNullOrEmpty(word) && !_ignoreWords.Contains(word))
            {
                lock (_ignoreWords)
                    _ignoreWords.Add(word);

                // Notify listeners.
                RaiseSpellingChangedEvent(word);
            }
        }

        public bool ShouldIgnoreWord(string word)
        {
            lock (_ignoreWords)
                return _ignoreWords.Contains(word);
        }

        public event EventHandler<SpellingEventArgs> DictionaryUpdated;

        #endregion

        #region Helpers

        void RaiseSpellingChangedEvent(string word)
        {
            var temp = DictionaryUpdated;
            if (temp != null)
                DictionaryUpdated(this, new SpellingEventArgs(word));
        }

        private void LoadIgnoreDictionary()
        {
            if (File.Exists(_ignoreWordsFile))
            {
                _ignoreWords.Clear();
                using (StreamReader reader = new StreamReader(_ignoreWordsFile))
                {
                    string word;
                    while (!string.IsNullOrEmpty((word = reader.ReadLine())))
                    {
                        _ignoreWords.Add(word);
                    }
                }
            }
        }
        #endregion
    }
}
