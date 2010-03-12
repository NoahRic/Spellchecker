using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Threading;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;
using SpellChecker.Definitions;

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
        ISpellingDictionaryService SpellingDictionaryFactory = null;

        public ITagger<T> CreateTagger<T>(ITextBuffer buffer) where T : ITag
        {
            var dictionary = SpellingDictionaryFactory.GetDictionary(buffer);

            return buffer.Properties.GetOrCreateSingletonProperty(
                () => new SpellingTagger(buffer, AggregatorFactory.CreateTagAggregator<INaturalTextTag>(buffer), dictionary)) as ITagger<T>;
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
        struct DirtySpan
        {
            public SnapshotSpan Span;
            public NormalizedSnapshotSpanCollection NaturalTextSpans;
        }

        ITextBuffer _buffer;
        ITagAggregator<INaturalTextTag> _naturalTextTagger;
        Dispatcher _dispatcher;
        ISpellingDictionary _dictionary;

        List<DirtySpan> _dirtySpans;

        object _dirtySpanLock = new object();
        volatile List<MisspellingTag> _misspellings;

        Thread _updateThread;

        DispatcherTimer _timer;

        public SpellingTagger(ITextBuffer buffer, ITagAggregator<INaturalTextTag> naturalTextTagger, ISpellingDictionary dictionary)
        {
            _buffer = buffer;
            _naturalTextTagger = naturalTextTagger;
            _dispatcher = Dispatcher.CurrentDispatcher;
            _dictionary = dictionary;

            _dirtySpans = new List<DirtySpan>();
            _misspellings = new List<MisspellingTag>();

            _buffer.Changed += BufferChanged;
            _naturalTextTagger.TagsChanged += NaturalTagsChanged;
            _dictionary.DictionaryUpdated += DictionaryUpdated;

            // To start with, the entire buffer is dirty
            // Split this into chunks, so we update pieces at a time
            ITextSnapshot snapshot = _buffer.CurrentSnapshot;

            foreach (var line in snapshot.Lines)
                AddDirtySpan(line.Extent);
        }

        void NaturalTagsChanged(object sender, TagsChangedEventArgs e)
        {
            NormalizedSnapshotSpanCollection dirtySpans = e.Span.GetSpans(_buffer.CurrentSnapshot);

            if (dirtySpans.Count == 0)
                return;

            SnapshotSpan dirtySpan = new SnapshotSpan(_buffer.CurrentSnapshot, dirtySpans[0].Start, dirtySpans[dirtySpans.Count - 1].End);

            AddDirtySpan(dirtySpan);

            var temp = TagsChanged;
            if (temp != null)
                temp(this, new SnapshotSpanEventArgs(dirtySpan));
        }

        void DictionaryUpdated(object sender, SpellingEventArgs e)
        {
            ITextSnapshot snapshot = _buffer.CurrentSnapshot;

            // If the word is null, it means the entire dictionary was updated and we
            // need to reparse the entire file.
            if (e.Word == null)
            {
                foreach (var line in snapshot.Lines)
                    AddDirtySpan(line.Extent);
                return;
            }

            List<MisspellingTag> currentMisspellings = _misspellings;

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

        NormalizedSnapshotSpanCollection GetNaturalLanguageSpansForDirtySpan(SnapshotSpan dirtySpan)
        {
            if (dirtySpan.IsEmpty)
                return new NormalizedSnapshotSpanCollection();

            ITextSnapshot snapshot = dirtySpan.Snapshot;
            return new NormalizedSnapshotSpanCollection(_naturalTextTagger.GetTags(dirtySpan)
                                                                          .SelectMany(tag => tag.Span.GetSpans(snapshot))
                                                                          .Select(s => s.Intersection(dirtySpan))
                                                                          .Where(s => s.HasValue && !s.Value.IsEmpty)
                                                                          .Select(s => s.Value));                    
        }

        void AddDirtySpan(SnapshotSpan span)
        {
            if (span.IsEmpty)
                return;

            var naturalLanguageSpans = GetNaturalLanguageSpansForDirtySpan(span);

            DirtySpan dirty = new DirtySpan() { NaturalTextSpans = naturalLanguageSpans, Span = span };

            lock (_dirtySpanLock)
            {
                _dirtySpans.Add(dirty);
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
                    // If an update is currently running, wait until the next timer tick
                    if (_updateThread != null && _updateThread.IsAlive)
                        return;

                    _timer.Stop();

                    _updateThread = new Thread(GuardedCheckSpellings)
                    {
                        Name = "Spell Check",
                        Priority = ThreadPriority.BelowNormal
                    };

                    if (!_updateThread.TrySetApartmentState(ApartmentState.STA))
                        Debug.Fail("Unable to set thread apartment state to STA, things *will* break.");

                    _updateThread.Start();
                };
            }

            _timer.Stop();
            _timer.Start();
        }

        void GuardedCheckSpellings(object obj)
        {
            try
            {
                CheckSpellings();
            }
            catch (Exception)
            {
                // If anything fails in the background thread, just ignore it.  It's possible that the background thread will run
                // on VS shutdown, at which point calls into WPF throw exceptions.  If we don't guard against those exceptions, the
                // user will see a crash on exit.
            }
        }

        void CheckSpellings()
        {
            TextBox textBox = new TextBox();
            textBox.SpellCheck.IsEnabled = true;

            IList<DirtySpan> dirtySpans;

            lock (_dirtySpanLock)
            {
                dirtySpans = _dirtySpans;
                if (dirtySpans.Count == 0)
                    return;

                _dirtySpans = new List<DirtySpan>();
            }

            ITextSnapshot snapshot = _buffer.CurrentSnapshot;

            // Break up dirty into component pieces, so we produce incremental updates
            foreach (var dirtySpan in dirtySpans)
            {
                var dirty = dirtySpan.Span.TranslateTo(snapshot, SpanTrackingMode.EdgeInclusive);
                var naturalText = new NormalizedSnapshotSpanCollection(
                    dirtySpan.NaturalTextSpans.Select(span => span.TranslateTo(snapshot, SpanTrackingMode.EdgeInclusive)));

                List<MisspellingTag> currentMisspellings = new List<MisspellingTag>(_misspellings);
                List<MisspellingTag> newMisspellings = new List<MisspellingTag>();

                int removed = currentMisspellings.RemoveAll(tag => tag.ToTagSpan(snapshot).Span.OverlapsWith(dirty));
                newMisspellings.AddRange(GetMisspellingsInSpans(naturalText, textBox));
               
                // Also remove empties
                removed += currentMisspellings.RemoveAll(tag => tag.ToTagSpan(snapshot).Span.IsEmpty);

                // If anything has been updated, we need to send out a change event
                if (newMisspellings.Count != 0 || removed != 0)
                {
                    currentMisspellings.AddRange(newMisspellings);

                    _dispatcher.Invoke(new Action(() =>
                        {
                            _misspellings = currentMisspellings;

                            var temp = TagsChanged;
                            if (temp != null)
                                temp(this, new SnapshotSpanEventArgs(dirty));

                        }));
                }
            }
            
            lock (_dirtySpanLock)
            {
                if (_dirtySpans.Count != 0)
                    _dispatcher.BeginInvoke(new Action(() => ScheduleUpdate()));
            }
        }
        
        IEnumerable<MisspellingTag> GetMisspellingsInSpans(NormalizedSnapshotSpanCollection spans, TextBox textBox)
        {
            foreach (var span in spans)
            {
                string text = span.GetText();

                // We need to break this up for WPF, because it is *incredibly* slow at checking the spelling
                for (int i = 0; i < text.Length; i++)
                {
                    if (text[i] == ' ' || text[i] == '\t' || text[i] == '\r' || text[i] == '\n')
                        continue;

                    // We've found a word (or something), so search for the next piece of whitespace or punctuation to get the entire word span.
                    // However, we will ignore words that are CamelCased, since those are probably not "real" words to begin with.
                    int end = i;
                    bool foundLower = false;
                    bool ignoreWord = false;
                    for (; end < text.Length; end++)
                    {
                        char c = text[end];

                        if (c == ' ' || c == '\t' || c == '\r' || c == '\n')
                            break;

                        if (!ignoreWord)
                        {
                            bool isUppercase = char.IsUpper(c);

                            if (foundLower && isUppercase)
                                ignoreWord = true;

                            foundLower = !isUppercase;
                        }
                    }

                    // Skip this word and move on to the next
                    if (ignoreWord)
                    {
                        i = end - 1;
                        continue;
                    }

                    string textToParse = text.Substring(i, end - i);

                    // Now pass these off to WPF
                    textBox.Text = textToParse;

                    int nextSearchIndex = 0;
                    int nextSpellingErrorIndex = -1;

                    while (-1 != (nextSpellingErrorIndex = textBox.GetNextSpellingErrorCharacterIndex(nextSearchIndex, LogicalDirection.Forward)))
                    {
                        var spellingError = textBox.GetSpellingError(nextSpellingErrorIndex);
                        int length = textBox.GetSpellingErrorLength(nextSpellingErrorIndex);

                        SnapshotSpan errorSpan = new SnapshotSpan(span.Snapshot, span.Start + i + nextSpellingErrorIndex, length);

                        if (!_dictionary.ShouldIgnoreWord(errorSpan.GetText()))
                        {
                            yield return new MisspellingTag(errorSpan, spellingError.Suggestions.ToArray());
                        }

                        nextSearchIndex = nextSpellingErrorIndex + length;
                        if (nextSearchIndex >= textToParse.Length)
                            break;
                    }

                    // Move past this word
                    i = end - 1;
                }
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
