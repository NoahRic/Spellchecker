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
using Microsoft.VisualStudio.Text.Editor;

namespace Microsoft.VisualStudio.Language.Spellchecker
{
    [Export(typeof(IViewTaggerProvider))]
    [ContentType("any")]
    [TagType(typeof(MisspellingTag))]
    sealed class SpellingTaggerProvider : IViewTaggerProvider
    {
        [Import]
        IViewTagAggregatorFactoryService AggregatorFactory = null;

        [Import]
        ISpellingDictionaryService SpellingDictionaryFactory = null;

        public ITagger<T> CreateTagger<T>(ITextView textView, ITextBuffer buffer) where T : ITag
        {
            if (textView.TextBuffer != buffer)
                return null;

            SpellingTagger spellingTagger;
            if (textView.Properties.TryGetProperty(typeof(SpellingTagger), out spellingTagger))
                return spellingTagger as ITagger<T>;

            var dictionary = SpellingDictionaryFactory.GetDictionary(buffer);
            var aggregator = AggregatorFactory.CreateTagAggregator<INaturalTextTag>(textView, TagAggregatorOptions.MapByContentType);
            spellingTagger = new SpellingTagger(buffer, aggregator, dictionary);
            textView.Properties[typeof(SpellingTagger)] = spellingTagger;

            return spellingTagger as ITagger<T>;
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

        public ITagSpan<MisspellingTag> ToTagSpan(ITextSnapshot snapshot)
        {
            return new TagSpan<MisspellingTag>(Span.GetSpan(snapshot), this);
        }
    }

    sealed class SpellingTagger : ITagger<MisspellingTag>
    {
        struct SpanToParse
        {
            public SnapshotSpan Span;
            public NormalizedSnapshotSpanCollection NaturalTextSpans;
        }

        ITextBuffer _buffer;
        ITagAggregator<INaturalTextTag> _naturalTextAggregator;
        Dispatcher _dispatcher;
        ISpellingDictionary _dictionary;

        List<SnapshotSpan> _dirtySpans;

        object _dirtySpanLock = new object();
        volatile List<MisspellingTag> _misspellings;

        Thread _updateThread;

        DispatcherTimer _timer;

        public SpellingTagger(ITextBuffer buffer, ITagAggregator<INaturalTextTag> naturalTextAggregator, ISpellingDictionary dictionary)
        {
            _buffer = buffer;
            _naturalTextAggregator = naturalTextAggregator;
            _dispatcher = Dispatcher.CurrentDispatcher;
            _dictionary = dictionary;

            _dirtySpans = new List<SnapshotSpan>();
            _misspellings = new List<MisspellingTag>();

            _buffer.Changed += BufferChanged;
            _naturalTextAggregator.TagsChanged += NaturalTagsChanged;
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

            foreach (var span in dirtySpans)
                AddDirtySpan(span);
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
            return new NormalizedSnapshotSpanCollection(_naturalTextAggregator.GetTags(dirtySpan)
                                                                              .SelectMany(tag => tag.Span.GetSpans(snapshot))
                                                                              .Select(s => s.Intersection(dirtySpan))
                                                                              .Where(s => s.HasValue && !s.Value.IsEmpty)
                                                                              .Select(s => s.Value));                    
        }

        void AddDirtySpan(SnapshotSpan span)
        {
            if (span.IsEmpty)
                return;

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

                _timer.Tick += StartUpdateThread;
            }

            _timer.Stop();
            _timer.Start();
        }

        void StartUpdateThread(object sender, EventArgs e)
        {
            // If an update is currently running, wait until the next timer tick
            if (_updateThread != null && _updateThread.IsAlive)
                return;

            _timer.Stop();

            List<SnapshotSpan> dirtySpans;
            lock (_dirtySpanLock)
            {
                dirtySpans = new List<SnapshotSpan>(_dirtySpans);
                _dirtySpans = new List<SnapshotSpan>();

                if (dirtySpans.Count == 0)
                    return;
            }

            // Normalize the dirty spans
            ITextSnapshot snapshot = _buffer.CurrentSnapshot;
            var normalizedSpans = new NormalizedSnapshotSpanCollection(dirtySpans.Select(s => s.TranslateTo(snapshot, SpanTrackingMode.EdgeInclusive)));
            var spansToParse = normalizedSpans.Select(s => new SpanToParse() { Span = s, NaturalTextSpans = GetNaturalLanguageSpansForDirtySpan(s) })
                                              .ToList();

            _updateThread = new Thread(GuardedCheckSpellings)
            {
                Name = "Spell Check",
                Priority = ThreadPriority.BelowNormal
            };

            if (!_updateThread.TrySetApartmentState(ApartmentState.STA))
                Debug.Fail("Unable to set thread apartment state to STA, things *will* break.");

            _updateThread.Start(spansToParse);
        }

        void GuardedCheckSpellings(object spansToParseObject)
        {
            try
            {
                IEnumerable<SpanToParse> spansToParse = spansToParseObject as IEnumerable<SpanToParse>;
                if (spansToParse == null)
                {
                    Debug.Fail("Being asked to check a null list of dirty spans.  What gives?");
                    return;
                }

                CheckSpellings(spansToParse);
            }
            catch (Exception ex)
            {
                Debug.Fail("Exception!" + ex.Message);
                // If anything fails in the background thread, just ignore it.  It's possible that the background thread will run
                // on VS shutdown, at which point calls into WPF throw exceptions.  If we don't guard against those exceptions, the
                // user will see a crash on exit.
            }
        }

        void CheckSpellings(IEnumerable<SpanToParse> spansToParse)
        {
            TextBox textBox = new TextBox();
            textBox.SpellCheck.IsEnabled = true;

            ITextSnapshot snapshot = _buffer.CurrentSnapshot;

            foreach (var spanToParse in spansToParse)
            {
                var dirty = spanToParse.Span.TranslateTo(snapshot, SpanTrackingMode.EdgeInclusive);
                var naturalText = new NormalizedSnapshotSpanCollection(
                    spanToParse.NaturalTextSpans.Select(span => span.TranslateTo(snapshot, SpanTrackingMode.EdgeInclusive)));

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
                    if (!IsSpellingWordChar(text[i]))
                        continue;

                    // We've found a word (or something), so search for the next piece of whitespace or punctuation to get the entire word span.
                    // However, we will ignore words in a few cases:
                    // 1) Words that are CamelCased, since those are probably not "real" words to begin with.
                    // 2) Things that look like filenames (contain a "." followed by something other than a "."). We may miss a few "real" misspellings
                    //    here due to a missed space after a period, but that's acceptable.
                    // 3) Words that include digits
                    // 4) Words that include underscores
                    int end = i;
                    bool foundLower = false;
                    bool ignoreWord = false;
                    bool lastLetterWasADot = false;

                    for (; end < text.Length; end++)
                    {
                        char c = text[end];

                        if (!ignoreWord)
                        {
                            bool isUppercase = char.IsUpper(c);

                            if (foundLower && isUppercase)
                                ignoreWord = true;
                            else if (c == '_')
                                ignoreWord = true;
                            else if (char.IsDigit(c))
                                ignoreWord = true;
                            else if (lastLetterWasADot && c != '.')
                                ignoreWord = true;

                            foundLower = char.IsLower(c);
                            lastLetterWasADot = (c == '.');
                        }

                        if (!IsSpellingWordChar(c))
                            break;
                    }

                    // If this word is in ALL CAPS, ignore it
                    if (!foundLower)
                        ignoreWord = true;

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

        /// <summary>
        /// Determine if the given character is a "spelling" word char, which includes a few more things than just characters
        /// </summary>
        bool IsSpellingWordChar(char c)
        {
            return c == '\'' || c == '`' || c == '-' || c == '.' || char.IsLetter(c);
        }

        #endregion

        #region Tagging implementation

        public IEnumerable<ITagSpan<MisspellingTag>> GetTags(NormalizedSnapshotSpanCollection spans)
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
