using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Text;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Utilities;
using System.Threading;
using System.Windows.Threading;
using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Documents;

namespace Microsoft.VisualStudio.Language.Spellchecker
{
    [Export(typeof(ITaggerProvider))]
    [ContentType("text")]
    [TagType(typeof(IMisspellingTag))]
    sealed class SpellingTaggerProvider : ITaggerProvider
    {
        [Import]
        IBufferTagAggregatorFactoryService AggregatorFactory = null;

        [Import]
        ISpellingDictionaryService SpellingDictionary = null;

        public ITagger<T> CreateTagger<T>(ITextBuffer buffer) where T : ITag
        {
            return buffer.Properties.GetOrCreateSingletonProperty(
                () => new SpellingTagger(buffer, AggregatorFactory.CreateTagAggregator<NaturalTextTag>(buffer), SpellingDictionary)) as ITagger<T>;
        }
    }

    class MisspellingTag : IMisspellingTag
    {
        public MisspellingTag(SnapshotSpan span, IEnumerable<string> suggestions)
        {
            Span = span.Snapshot.CreateTrackingSpan(span, SpanTrackingMode.EdgeExclusive);
            Suggestions = suggestions;
        }

        public ITrackingSpan Span { get; private set; }
        public IEnumerable<string> Suggestions { get; private set; }

        public ITagSpan<IMisspellingTag> ToTagSpan(ITextSnapshot snapshot)
        {
            return new TagSpan<IMisspellingTag>(Span.GetSpan(snapshot), this);
        }
    }

    sealed class SpellingTagger : ITagger<IMisspellingTag>
    {
        ITextBuffer _buffer;
        ITagAggregator<NaturalTextTag> _naturalTextTagger;
        Dispatcher _dispatcher;
        ISpellingDictionaryService _dictionary;
        
        List<SnapshotSpan> _dirtySpans;
        object _dirtySpanLock = new object();
        volatile List<MisspellingTag> _misspellings;

        DispatcherTimer _timer;

        public SpellingTagger(ITextBuffer buffer, ITagAggregator<NaturalTextTag> naturalTextTagger, ISpellingDictionaryService dictionary)
        {
            _buffer = buffer;
            _naturalTextTagger = naturalTextTagger;
            _dispatcher = Dispatcher.CurrentDispatcher;
            _dictionary = dictionary;

            _dirtySpans = new List<SnapshotSpan>();
            _misspellings = new List<MisspellingTag>();

            _buffer.Changed += BufferChanged;

            _dictionary.DictionaryUpdated += DictionaryUpdated;

            // To start with, the entire buffer is dirty
            // Split this into chunks, so we update pieces at a time
            ITextSnapshot snapshot = _buffer.CurrentSnapshot;

            foreach (var line in snapshot.Lines)
                AddDirtySpan(line.Extent);
        }

        void DictionaryUpdated(object sender, SpellingEventArgs e)
        {
            List<MisspellingTag> currentMisspellings = _misspellings;
            ITextSnapshot snapshot = _buffer.CurrentSnapshot;

            foreach (var misspelling in currentMisspellings)
            {
                SnapshotSpan span = misspelling.Span.GetSpan(snapshot);

                if (span.GetText() == e.Word)
                    AddDirtySpan(span);
            }
        }

        void BufferChanged(object sender, TextContentChangedEventArgs e)
        {
            ITextSnapshot snapshot = e.After;

            foreach (var change in e.Changes)
            {
                SnapshotSpan changedSpan = new SnapshotSpan(snapshot, change.NewSpan);

                var startLine = changedSpan.Start.GetContainingLine();
                var endLine = (startLine.EndIncludingLineBreak < changedSpan.End) ? changedSpan.End.GetContainingLine() : startLine;

                AddDirtySpan(new SnapshotSpan(startLine.Start, endLine.End));
            }
        }

        #region Helpers


        void AddDirtySpan(SnapshotSpan span)
        {
            lock (_dirtySpanLock)
            {
                _dirtySpans.Add(span);
                ScheduleUpdate();
            }
        }

        void ScheduleUpdate()
        {
            if (_timer == null)
            {
                _timer = new DispatcherTimer(DispatcherPriority.ApplicationIdle, _dispatcher)
                {
                    Interval = TimeSpan.FromMilliseconds(500)
                };

                _timer.Tick += (sender, args) =>
                {
                    _timer.Stop();

                    Thread thread = new Thread(CheckSpellings)
                    {
                        Name = "Spell Check",
                        Priority = ThreadPriority.BelowNormal
                    };

                    if (!thread.TrySetApartmentState(ApartmentState.STA))
                        Debug.Fail("Unable to set thread apartment state to STA, things *will* break.");

                    thread.Start();
                };
            }

            _timer.Stop();
            _timer.Start();
        }

        void CheckSpellings(object obj)
        {
            TextBox textBox = new TextBox();
            textBox.SpellCheck.IsEnabled = true;

            IList<SnapshotSpan> dirtySpans;

            lock (_dirtySpanLock)
            {
                dirtySpans = _dirtySpans;
                if (dirtySpans.Count == 0)
                    return;

                _dirtySpans = new List<SnapshotSpan>();
            }

            ITextSnapshot snapshot = _buffer.CurrentSnapshot;

            NormalizedSnapshotSpanCollection dirty = new NormalizedSnapshotSpanCollection(
                dirtySpans.Select(span => span.TranslateTo(snapshot, SpanTrackingMode.EdgeInclusive)));

            if (dirty.Count == 0)
            {
                Debug.Fail("The list of dirty spans is empty when normalized, which shouldn't be possible.");
                return;
            }

            // Break up dirty into component pieces, so we produce incremental updates
            foreach (var dirtySpan in dirty)
            {
                List<MisspellingTag> currentMisspellings = _misspellings;
                List<MisspellingTag> newMisspellings = new List<MisspellingTag>();

                int removed = 0;

                foreach (var naturalTextTag in _naturalTextTagger.GetTags(dirtySpan))
                {
                    var naturalTextSpans = naturalTextTag.Span.GetSpans(snapshot);

                    foreach (var span in naturalTextSpans)
                    {
                        removed += currentMisspellings.RemoveAll(tag => tag.ToTagSpan(snapshot).Span.OverlapsWith(span));

                        newMisspellings.AddRange(GetMisspellingsInSpan(span, textBox));
                    }
                }

                // Also remove empties
                removed += currentMisspellings.RemoveAll(tag => tag.ToTagSpan(snapshot).Span.Length == 0);

                // If anything has been updated, we need to send out a change event
                if (newMisspellings.Count != 0 || removed != 0)
                {
                    currentMisspellings.AddRange(newMisspellings);
                    SnapshotSpan changedSpan = new SnapshotSpan(dirty[0].Start, dirty[dirty.Count - 1].End);

                    _dispatcher.Invoke(new Action(() =>
                        {
                            _misspellings = currentMisspellings;

                            var temp = TagsChanged;
                            if (temp != null)
                                temp(this, new SnapshotSpanEventArgs(changedSpan));

                        }));
                }
            }
            
            lock (_dirtySpanLock)
            {
                if (_dirtySpans.Count != 0)
                    _dispatcher.Invoke(new Action(() => ScheduleUpdate()));
            }
        }
        
        IEnumerable<MisspellingTag> GetMisspellingsInSpan(SnapshotSpan span, TextBox textBox)
        {
            string text = span.GetText();

            // We need to break this up for WPF, because it is *incredibly* slow at checking the spelling
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == ' ' || text[i] == '\t' || text[i] == '\r' || text[i] == '\n')
                    continue;

                // We've found a word (or something)
                // Scan until the next whitespace
                int end = i;
                for (; end < text.Length; end++)
                {
                    if (text[end] == ' ' || text[end] == '\t' || text[end] == '\r' || text[end] == '\n')
                        break;
                }

                string wordsToParse = text.Substring(i, end - i);

                // Now pass these off to WPF
                textBox.Text = wordsToParse;

                int nextSearchIndex = 0;
                int nextSpellingErrorIndex = -1;

                while (-1 != (nextSpellingErrorIndex = textBox.GetNextSpellingErrorCharacterIndex(nextSearchIndex, LogicalDirection.Forward)))
                {
                    var spellingError = textBox.GetSpellingError(nextSpellingErrorIndex);
                    int length = textBox.GetSpellingErrorLength(nextSpellingErrorIndex);

                    SnapshotSpan errorSpan = new SnapshotSpan(span.Snapshot, span.Start + i + nextSpellingErrorIndex, length);

                    if (_dictionary.IsWordInDictionary(errorSpan.GetText()))
                    {
                        spellingError.IgnoreAll();
                    }
                    else
                    {
                        yield return new MisspellingTag(errorSpan, spellingError.Suggestions.ToArray());
                    }

                    nextSearchIndex = nextSpellingErrorIndex + length;
                    if (nextSearchIndex >= wordsToParse.Length)
                        break;
                }

                // Move past this word
                i = end - 1;
            }
        }

        #endregion

        #region Tagging implementation

        public IEnumerable<ITagSpan<IMisspellingTag>> GetTags(NormalizedSnapshotSpanCollection spans)
        {
            if (spans.Count == 0)
                yield break;

            List<MisspellingTag> currentMisspellings = _misspellings;

            if (currentMisspellings.Count == 0)
                yield break;

            ITextSnapshot snapshot = spans[0].Snapshot;

            foreach (var misspelling in currentMisspellings)
            {
                var tagSpan = misspelling.ToTagSpan(snapshot);
                if (tagSpan.Span.Length == 0)
                    continue;

                if (spans.IntersectsWith(new NormalizedSnapshotSpanCollection(tagSpan.Span)))
                    yield return tagSpan;
            }
        }

        public event EventHandler<SnapshotSpanEventArgs> TagsChanged;
        
        #endregion
    }
}
