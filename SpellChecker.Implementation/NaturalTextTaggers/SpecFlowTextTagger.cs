using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;

namespace Microsoft.VisualStudio.Language.Spellchecker
{
    /// <summary>
    /// Provides tags for SpecFlow / Gherkin Feature files.
    /// </summary>
    internal class SpecFlowTextTagger : ITagger<NaturalTextTag>
    {
        #region Private Fields
        private ITextBuffer _buffer;
        #endregion

        #region MEF Imports / Exports
        [Export(typeof(ITaggerProvider))]
        [ContentType("gherkin")]
        [FileExtension(".feature")]
        [TagType(typeof(NaturalTextTag))]
        internal class SpecFlowTaggerProvider : ITaggerProvider
        {
            public ITagger<T> CreateTagger<T>(ITextBuffer buffer) where T : ITag
            {
                if (buffer == null)
                {
                    throw new ArgumentNullException("buffer");
                }

                return new SpecFlowTextTagger(buffer) as ITagger<T>;
            }
        }
        #endregion

        #region Constructor

        public SpecFlowTextTagger(ITextBuffer buffer)
        {
            _buffer = buffer;
        }
        #endregion

        #region ITagger<INaturalTextTag> Members
        public IEnumerable<ITagSpan<NaturalTextTag>> GetTags(NormalizedSnapshotSpanCollection spans)
        {
            foreach (var snapshotSpan in spans)
            {
                yield return new TagSpan<NaturalTextTag>(
                        snapshotSpan,
                        new NaturalTextTag()
                        );

            }
        }

        public event EventHandler<SnapshotSpanEventArgs> TagsChanged
        {
            add { }
            remove { }
		  }
		  #endregion
	 }
}
