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
using System.Diagnostics;
using System.Globalization;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;
using System.Linq;

namespace Microsoft.VisualStudio.Language.Spellchecker
{
    [Export(typeof(ITaggerProvider))]
    [ContentType("code")]
    [TagType(typeof(NaturalTextTag))]
    internal class CommentTextTaggerProvider : ITaggerProvider
    {
        [Import]
        internal IClassifierAggregatorService ClassifierAggregatorService { get; set; }

        [Import]
        internal IBufferTagAggregatorFactoryService TagAggregatorFactory { get; set; }

        public ITagger<T> CreateTagger<T>(ITextBuffer buffer) where T : ITag
        {
            if (buffer == null)
                throw new ArgumentNullException("buffer");

            var urlAggregator = TagAggregatorFactory.CreateTagAggregator<IUrlTag>(buffer);
            var classifierAggregator = ClassifierAggregatorService.GetClassifier(buffer);

            return new CommentTextTagger(buffer, classifierAggregator, urlAggregator) as ITagger<T>;
        }
    }

    internal class CommentTextTagger : ITagger<NaturalTextTag>, IDisposable
    {
        ITextBuffer _buffer;
        IClassifier _classifier;
        ITagAggregator<IUrlTag> _urlAggregator;

        public CommentTextTagger(ITextBuffer buffer, IClassifier classifier, ITagAggregator<IUrlTag> urlAggregator)
        {
            _buffer = buffer;
            _classifier = classifier;
            _urlAggregator = urlAggregator;

            classifier.ClassificationChanged += ClassificationChanged;
            urlAggregator.TagsChanged += UrlTagsChanged;
        }

        public IEnumerable<ITagSpan<NaturalTextTag>> GetTags(NormalizedSnapshotSpanCollection spans)
        {
            if (_classifier == null || spans == null || spans.Count == 0)
                yield break;

            ITextSnapshot snapshot = spans[0].Snapshot;

            // First, subtract out any URLs
            var urlSpans = new NormalizedSnapshotSpanCollection(
                _urlAggregator.GetTags(spans)
                              .SelectMany(tagSpan => tagSpan.Span.GetSpans(snapshot)));

            spans = NormalizedSnapshotSpanCollection.Difference(spans, urlSpans);

            // Now, search the URLs for human-readable text
            foreach (var snapshotSpan in spans)
            {
                Debug.Assert(snapshotSpan.Snapshot.TextBuffer == _buffer);
                foreach (ClassificationSpan classificationSpan in _classifier.GetClassificationSpans(snapshotSpan))
                {
                    string name = classificationSpan.ClassificationType.Classification.ToLowerInvariant();

                    if ((name.Contains("comment") || name.Contains("string")) &&
                       !(name.Contains("xml doc tag")))
                    {
                        yield return new TagSpan<NaturalTextTag>(
                                classificationSpan.Span,
                                new NaturalTextTag()
                                );
                    }
                }
            }
        }

        void ClassificationChanged(object sender, ClassificationChangedEventArgs e)
        {
            var temp = TagsChanged;
            if (temp != null)
                temp(this, new SnapshotSpanEventArgs(e.ChangeSpan));
        }

        void UrlTagsChanged(object sender, TagsChangedEventArgs e)
        {
            var temp = TagsChanged;
            if (temp != null)
            {
                var spans = e.Span.GetSpans(_buffer.CurrentSnapshot);
                if (spans.Count == 0)
                    return;

                temp(this, new SnapshotSpanEventArgs(new SnapshotSpan(spans[0].Start, spans[spans.Count - 1].End)));
            }
        }

        public event EventHandler<SnapshotSpanEventArgs> TagsChanged;

        public void Dispose()
        {
            if (_urlAggregator != null)
                _urlAggregator.TagsChanged -= UrlTagsChanged;
            if (_classifier != null)
                _classifier.ClassificationChanged -= ClassificationChanged;
        }
    }
}
