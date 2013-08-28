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

using System.Collections.ObjectModel;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using SpellChecker.Definitions;
using System.Diagnostics;

namespace Microsoft.VisualStudio.Language.Spellchecker
{
    /// <summary>
    /// Smart tag action for adding new words to the dictionary.
    /// </summary>
    internal class SpellAddRemoveLanguageSmartTagAction : ISmartTagAction
    {

        #region Constructor
        /// <summary>
        /// Constructor for SpellAddRemoveLanguageSmartTagAction.
        /// </summary>
        /// <param name="word">The word to add or ignore.</param>
        /// <param name="dictionary">The dictionary (used to ignore the word).</param>
        /// <param name="displayText">Text to show in the context menu for this action.</param>
        /// <param name="ignore">Whether this is to ignore the word or add it to the dictionary.</param>
        public SpellAddRemoveLanguageSmartTagAction()
        {
            DisplayText = "Add or Remove Languages...";
        }
        # endregion

        #region ISmartTagAction implementation
        /// <summary>
        /// Text to display in the context menu.
        /// </summary>
        public string DisplayText
        {
            get;
            private set;
        }

        /// <summary>
        /// Icon to place next to the display text.
        /// </summary>
        public System.Windows.Media.ImageSource Icon
        {
            get { return null; }
        }

        /// <summary>
        /// This method is executed when action is selected in the context menu.
        /// </summary>
        public void Invoke()
        {
            var window = new AddRemoveWindow();
            window.Width = 500;
            window.Height = 600;
            window.Show();
        }

        /// <summary>
        /// Enable/disable this action.
        /// </summary>
        public bool IsEnabled
        {
            get { return true; }
        }

        /// <summary>
        /// Action set to make sub menus.
        /// </summary>
        public ReadOnlyCollection<SmartTagActionSet> ActionSets
        {
            get
            {
                return null;
            }
        }
        #endregion
    }
}
