using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Globalization;
using System.Reflection;
using System.IO;

namespace Microsoft.VisualStudio.Language.Spellchecker {

	/// <summary>
	/// Interaction logic for UserControl1.xaml
	/// </summary>
	public partial class AddRemoveWindow: Window {
		public AddRemoveWindow() {
			InitializeComponent();

			var cultures = CultureInfo.GetCultures(CultureTypes.AllCultures)
				.OrderBy(c => c.EnglishName);
			foreach (var culture in cultures) cultureSelector.Items.Add(culture);

			save.Click += (sender, args) => {
				Infos.Save();
				Close();
			};

			cancel.Click += (sender, args) => {
				Close();
			};

			add.Click += AddCulture;

			cultureSelector.SelectedItem = System.Threading.Thread.CurrentThread.CurrentCulture;
			cultureSelector.SelectionChanged += AddCulture;

			Infos.Apply();
		}

		public void AddCulture(object sender, EventArgs args) {
			var info = Infos[((CultureInfo)cultureSelector.SelectedValue).Name];
			info.Enabled = false;
			
			// auto import .dic & .lex dictionaries if available.
			var uri = SpellingTagger.ImportDefaultDictionaries(info.Culture.Name);
			if (uri != null) {
				info.CustomDictionaries.Add("*");
			}

			Infos.Apply();
			Infos.Save();
		}

		public class Info {
			string[] Native = new string[] { "en-US", "es-ES", "de-DE", "fr-FR" };
			public bool Enabled { get; set; }
			public CultureInfo Culture { get; set; }
			public List<string> CustomDictionaries { get; set; }
			public bool IsNative { get { return Native.Any(c => c == Culture.Name); } }
		}

		public class InfoCollection : KeyedCollection<string, Info> {

			public InfoCollection(AddRemoveWindow window) { Window = window; }

			AddRemoveWindow Window;

			protected override string GetKeyForItem(Info item) { return item.Culture.Name; }

			public new Info this[string culture] {
				get {
					if (!Contains(culture)) base.Add(new Info { Culture = new CultureInfo(culture), CustomDictionaries = new List<string>() });
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
					SpellingTagger.Languages = sb.ToString();
				}
			}

			public Image Image(string source, string toolTip = null) {
				var img = new Image();
				var src = new BitmapImage();
				src.BeginInit();
				src.UriSource = new Uri("pack://application:,,,/SpellChecker;Component/" + source);
				src.CacheOption = BitmapCacheOption.OnLoad;
				src.EndInit();
				img.Source = src;
				img.Stretch = Stretch.Uniform;
				img.ToolTip = toolTip;
				return img;
			}

			public void Apply() {

				Window.panel.Children.RemoveRange(0, Window.panel.Children.Count-1);

				int i = 0;
				foreach (var info in this.OrderBy(e => e.Culture.EnglishName)) {
					var index = i++;
					var checkbox = new CheckBox();
					checkbox.IsChecked = info.Enabled;
					checkbox.IsThreeState = false;
					checkbox.Content = info.Culture.EnglishName;
					var delete = new Button();
					delete.Style = (Style)Window.panel.Resources["ImageButton"];
					delete.Content = Image("Resources/cancel.gif", "Delete this language entry.");
					delete.Margin = new Thickness(0, 0, 8, 0);

					var stack = new StackPanel();
					stack.Orientation = Orientation.Horizontal;
					stack.Children.Add(delete);
					stack.Children.Add(checkbox);

					var defaultDict = new System.Windows.Controls.Primitives.ToggleButton();
					defaultDict.IsThreeState = false;
					defaultDict.Content = Image("Resources/anchor.png", "Use the standard dictionary.");

					var customDict = new System.Windows.Controls.Primitives.ToggleButton();
					customDict.IsThreeState = false;
					customDict.Content = Image("Resources/book_open.png", "Use custom dictionaries.");

					var rightstack = new StackPanel();
					rightstack.Orientation = Orientation.Horizontal;
					rightstack.HorizontalAlignment = HorizontalAlignment.Right;
					rightstack.Children.Add(defaultDict);
					rightstack.Children.Add(customDict);

					var dock = new DockPanel();
					stack.SetValue(DockPanel.DockProperty, Dock.Left);
					rightstack.SetValue(DockPanel.DockProperty, Dock.Right);
					dock.Children.Add(stack);
					dock.Children.Add(rightstack);
					dock.Background = new SolidColorBrush(Color.FromRgb(0xff, 0xff, 0xff));
					dock.Opacity = 0.88;

					checkbox.Click += (sender, args) => {
						info.Enabled = checkbox.IsChecked ?? false;
					};

					delete.Click += (sender, args) => {
						this.Remove(info);
						Window.panel.Children.Remove(dock);
					};

					defaultDict.Click += (sender, args) => {
						var has = info.CustomDictionaries.Any(d => d == "*");
 						if (has != (defaultDict.IsChecked ?? false)) {
							if (!has) info.CustomDictionaries.Add("*");
							else info.CustomDictionaries.Remove("*");
						}
					};

					customDict.Click += (sender, args) => {
						if (customDict.IsChecked ?? false) {
							var d = new System.Windows.Forms.OpenFileDialog();
							d.Title = "Select Custom Dictionary Files for " + info.Culture.EnglishName;
							d.SupportMultiDottedExtensions = true;
							d.Multiselect = true;
							d.Filter = "Dictionary Files (*.lex;*.txt;*.dic)|*.lex;*.txt;*.dic|All files (*.*)|*.*";
							d.ShowDialog();
							foreach (var file in d.FileNames) {
								var newfile = System.IO.Path.Combine(SpellingTagger.ConfigDirectory, System.IO.Path.GetFileName(file));
								var ext = System.IO.Path.GetExtension(file);
								if (ext == ".dic") { // import dic files from NetSpell & ISpell
									SpellingTagger.ImportDic(file);
								} else {
									System.IO.File.Copy(file, newfile);
								}
								info.CustomDictionaries.Add(System.IO.Path.GetFileName(newfile));
							}
						} else {
  							foreach (var customd in info.CustomDictionaries.Where(d => d != "*").ToList()) {
								var file = System.IO.Path.Combine(SpellingTagger.ConfigDirectory, customd);
								if (File.Exists(file)) File.Delete(file);
								info.CustomDictionaries.Remove(customd);
							}
						}
					};

					Window.panel.Children.Insert(index, dock);
				}
			}
		}

		InfoCollection infos = null;
		public InfoCollection Infos {
			get {
				if (infos == null) {
					infos = new InfoCollection(this);
					var langs = SpellingTagger.Languages
						.Split(new char[] { ';', ',' }, System.StringSplitOptions.RemoveEmptyEntries);
					foreach (var lang in langs) {
						var tokens = lang.Split(':');
						var culture = tokens.First().Trim();
						var enabled = culture[0] != '#';
						if (!enabled) culture = culture.Substring(1);
						var info = new Info { Culture = new CultureInfo(culture), Enabled = enabled, CustomDictionaries = tokens.Skip(1).ToList() };
						infos.Add(info);
					}
				}
				return infos;
			}
		}
	}
}
