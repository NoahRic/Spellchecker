using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Reflection;
using System.IO;

namespace Microsoft.VisualStudio.Language.Spellchecker {

	public class Configuration {

		static string configDir = null;
		public static string ConfigDirectory {
			get {
				if (configDir == null) {
					var info = new System.IO.DirectoryInfo(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VisualStudioSpellChecker"));
					if (!info.Exists) info.Create();
					configDir = info.FullName;
				}
				return configDir;
			}
		}

		const string DefaultLanguages = "en;#de;#fr;#es"; // Insert custom default dictionaries from the Dictionaries folder that are included in the vsix in the dll path here. See syntax of the Languages Property.

		static string languagesstring = null;
		// syntax:[#] culture { : dictionary-file } { ; [#] culture { : dictionary-file } } ), where dictionary-file is a filename of a dictionary file located in the ConfigDirectory. # entries are available but disabled.
		public static string LanguagesString {
			get {
				if (languagesstring == null) {
					var file = Path.Combine(ConfigDirectory, "languages.config");
					if (!File.Exists(file)) languagesstring = DefaultLanguages;
					else languagesstring = File.ReadAllLines(file).FirstOrDefault() ?? DefaultLanguages;
				}
				return languagesstring;
			}
			set {
				languagesstring = value;
				var file = Path.Combine(ConfigDirectory, "languages.config");
				File.WriteAllText(file, languagesstring);
			}
		}


		public static string ImportDic(string file) { // import dic files from NetSpell & ISpell
			var newfile = Path.Combine(ConfigDirectory, Path.GetFileName(file));
			newfile = Path.ChangeExtension(newfile, "lex");
			var ext = Path.GetExtension(file);
			if (ext == ".dic" && !File.Exists(newfile)) {
				var lines = File.ReadAllLines(file)
					.SkipWhile(s => s != "[Words]")
					.Skip(1)
					.Select(s => s.Split('/').First());
				File.WriteAllLines(newfile, lines);
			}
			return newfile;
		}

		public static Uri ImportDefaultDictionaries(string culture) { // import default .dic & .lex files from config & dll directory.
			var id = "_" + culture;
			var configid = Path.Combine(ConfigDirectory, id);
			var dllid = Path.Combine(Assembly.GetExecutingAssembly().CodeBase, id);

			if (!File.Exists(configid + ".lex")) {
				if (File.Exists(configid + ".dic")) ImportDic(configid + ".dic");
				if (File.Exists(dllid + ".dic")) ImportDic(dllid + ".dic");
				if (File.Exists(dllid + ".lex")) File.Copy(dllid + ".lex", configid + ".lex");
			}
			if (File.Exists(configid + ".lex")) return new Uri(configid + ".lex");
			else return null;
		}

		public static Uri GetDictUri(Language lang, string dict) {
			if (dict == "*") return Configuration.ImportDefaultDictionaries(lang.Culture.Name);
			else return new Uri("file://" + Configuration.ConfigDirectory.Replace("\\", "/") + "/" + dict);
		}

		public class Language {
			string[] Native = new string[] { "en", "en-US", "es", "es-ES", "de", "de-DE", "fr", "fr-FR" };
			public bool Enabled { get; set; }
			public CultureInfo Culture { get; set; }
			public List<string> CustomDictionaries { get; set; }
			public bool IsNative { get { return Native.Any(c => c == Culture.Name); } }
			public string StandardDictionary { get { return Path.Combine(ConfigDirectory, "_" + Culture.Name + ".lex"); } }
			public bool HasStandardDictionary { get { return File.Exists(StandardDictionary); } }
		}

		public class LanguageCollection : KeyedCollection<string, Language> {

			public LanguageCollection() { }

			protected override string GetKeyForItem(Language item) { return item.Culture.Name; }

			public new Language this[string culture] {
				get {
					if (!Contains(culture)) base.Add(new Language { Culture = new CultureInfo(culture), CustomDictionaries = new List<string>() });
					return base[culture];
				}
				set {
					if (value.Culture.Name != culture) value.Culture = new CultureInfo(culture);
					base.Add(value);
				}
			}

			public void Save() {
				StringBuilder sb = new StringBuilder();

				int n = 0;
				foreach (var info in this.OrderBy(e => e.Culture.EnglishName)) {
					if (n++ > 0) sb.Append(';');

					if (!info.Enabled) sb.Append('#');
					sb.Append(info.Culture.Name);
					foreach (var customDict in info.CustomDictionaries) {
						sb.Append(':');
						sb.Append(customDict);
					}
					LanguagesString = sb.ToString();
				}
			}
		}

		static LanguageCollection languages = null;
		public static LanguageCollection Languages {
			get {
				if (languages == null) {
					languages = new LanguageCollection();
					var langs = LanguagesString
						.Split(new char[] { ';', ',' }, System.StringSplitOptions.RemoveEmptyEntries);
					foreach (var lang in langs) {
						var tokens = lang.Split(':');
						var culture = tokens.First().Trim();
						var enabled = culture[0] != '#';
						if (!enabled) culture = culture.Substring(1);
						var info = new Language { Culture = new CultureInfo(culture), Enabled = enabled, CustomDictionaries = tokens.Skip(1).ToList() };
						languages.Add(info);
					}
				}
				return languages;
			}
		}
	}
}