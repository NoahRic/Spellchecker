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

using System.Collections.Generic;
using System.Linq;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using SpellChecker.Definitions;
using System.Diagnostics;

namespace Microsoft.VisualStudio.Language.Spellchecker
{
    /// <summary>
    /// Smart tag action for language settings.
    /// </summary>
    internal class SpellLanguageSettingsSmartTagAction : ISmartTagAction
    {

        #region Constructor
        /// <summary>
        /// Constructor for SpellLanguageSettingsSmartTagAction.
        /// </summary>
        public SpellLanguageSettingsSmartTagAction()
        {
            DisplayText = "Language Settings...";
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
        public void Invoke() {
		
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
				var langs = Configuration.Languages
					.Select(lang => new SpellLanguageSmartTagItem(lang.Culture.Name))
					.ToList<ISmartTagAction>();

				var addremove = new List<ISmartTagAction>();
				addremove.Add(new SpellAddRemoveLanguageSmartTagAction());

				var sets = new List<SmartTagActionSet>();
				sets.Add(new SmartTagActionSet(langs.AsReadOnly()));
				sets.Add(new SmartTagActionSet(addremove.AsReadOnly()));

				return sets.AsReadOnly();

            }
        }
        #endregion
    }
}
