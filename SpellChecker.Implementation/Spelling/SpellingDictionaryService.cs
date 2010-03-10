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
using SpellChecker.Definitions;

namespace Microsoft.VisualStudio.Language.Spellchecker
{
    /// <summary>
    /// Spell checking provider based on WPF spell checker
    /// </summary>
    [Export(typeof(ISpellingDictionaryService))]
    internal class SpellingDictionaryServiceFactory : ISpellingDictionaryService
    {
        [ImportMany(typeof(IBufferSpecificDictionaryProvider))]
        IEnumerable<Lazy<IBufferSpecificDictionaryProvider>> BufferSpecificDictionaryProviders = null;

        public ISpellingDictionary GetDictionary(ITextBuffer buffer)
        {
            ISpellingDictionary service;
            if (buffer.Properties.TryGetProperty(typeof(SpellingDictionaryService), out service))
                return service;

            List<ISpellingDictionary> bufferSpecificDictionaries = new List<ISpellingDictionary>();

            foreach (var provider in BufferSpecificDictionaryProviders)
            {
                var dictionary = provider.Value.GetDictionary(buffer);
                if (dictionary != null)
                    bufferSpecificDictionaries.Add(dictionary);
            }

            service = new SpellingDictionaryService(bufferSpecificDictionaries);
            buffer.Properties[typeof(SpellingDictionaryService)] = service;

            return service;
        }
    }

    internal class SpellingDictionaryService : ISpellingDictionary
    {
        #region Private data
        private SortedSet<string> _ignoreWords = new SortedSet<string>();
        private IList<ISpellingDictionary> _bufferSpecificDictionaries;
        private string _ignoreWordsFile;
        private bool _gotBufferSpecificEvent;
        #endregion

        #region Constructor
        /// <summary>
        /// Constructor for SpellingDictionaryService
        /// </summary>
        public SpellingDictionaryService(IList<ISpellingDictionary> bufferSpecificDictionaries)
        {
            _bufferSpecificDictionaries = bufferSpecificDictionaries;

            foreach (var dictionary in _bufferSpecificDictionaries)
            {
                dictionary.DictionaryUpdated += BufferSpecificDictionaryUpdated;
            }

            string localFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\VisualStudio\10.0\SpellChecker");
            if (!Directory.Exists(localFolder))
            {
                Directory.CreateDirectory(localFolder);
            }
            _ignoreWordsFile = Path.Combine(localFolder, "Dictionary.txt");

            LoadIgnoreDictionary();
        }

        void BufferSpecificDictionaryUpdated(object sender, SpellingEventArgs e)
        {
            _gotBufferSpecificEvent = true;
            RaiseSpellingChangedEvent(e.Word);
        }
        #endregion

        #region ISpellingDictionaryService

        /// <summary>
        /// Adds given word to the dictionary.
        /// </summary>
        /// <param name="word">The word to add to the dictionary.</param>
        public bool AddWordToDictionary(string word)
        {
            if (!string.IsNullOrEmpty(word))
            {
                _gotBufferSpecificEvent = false;

                foreach (var dictionary in _bufferSpecificDictionaries)
                {
                    if (dictionary.AddWordToDictionary(word))
                    {
                        if (!_gotBufferSpecificEvent)
                            RaiseSpellingChangedEvent(word);
                        return true;
                    }
                }

                // Add this word to the dictionary file.
                using (StreamWriter writer = new StreamWriter(_ignoreWordsFile, true))
                {
                    writer.WriteLine(word);
                }

                IgnoreWord(word, addedToDictionary: true);

                return true;
            }

            return false;
        }

        public bool IgnoreWord(string word)
        {
            return IgnoreWord(word, false);
        }

        private bool IgnoreWord(string word, bool addedToDictionary)
        {
            if (!string.IsNullOrEmpty(word) && !_ignoreWords.Contains(word))
            {
                if (!addedToDictionary)
                {
                    _gotBufferSpecificEvent = false;

                    foreach (var dictionary in _bufferSpecificDictionaries)
                    {
                        if (dictionary.IgnoreWord(word))
                        {
                            if (!_gotBufferSpecificEvent)
                                RaiseSpellingChangedEvent(word);
                            return true;
                        }
                    }
                }

                lock (_ignoreWords)
                    _ignoreWords.Add(word);

                // Notify listeners.
                RaiseSpellingChangedEvent(word);

                return true;
            }

            return false;
        }

        public bool ShouldIgnoreWord(string word)
        {
            foreach (var dictionary in _bufferSpecificDictionaries)
            {
                if (dictionary.ShouldIgnoreWord(word))
                    return true;
            }

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
