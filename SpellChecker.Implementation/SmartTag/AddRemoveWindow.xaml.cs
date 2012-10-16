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
				Configuration.Languages.Save();
				Close();
			};

			cancel.Click += (sender, args) => {
				Close();
			};

			add.Click += AddCulture;

			cultureSelector.SelectedItem = System.Threading.Thread.CurrentThread.CurrentCulture;
			cultureSelector.SelectionChanged += AddCulture;

			Apply();
		}

		public void AddCulture(object sender, EventArgs args) {
			var info = Configuration.Languages[((CultureInfo)cultureSelector.SelectedValue).Name];
			info.Enabled = false;
			
			// auto import .dic & .lex dictionaries if available.
			var uri = Configuration.ImportDefaultDictionaries(info.Culture.Name);
			if (uri != null) {
				info.CustomDictionaries.Add("*");
			}

			Apply();
			Configuration.Languages.Save();
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

			panel.Children.RemoveRange(0, panel.Children.Count-1);

			int i = 0;
			foreach (var language in Configuration.Languages.OrderBy(e => e.Culture.EnglishName)) {
				var lang = language;
				var index = i++;
				var checkbox = new CheckBox();
				checkbox.IsChecked = lang.Enabled;
				checkbox.IsThreeState = false;
				checkbox.Content = lang.Culture.EnglishName;
				var delete = new Button();
				delete.Style = (Style)panel.Resources["ImageButton"];
				delete.Content = Image("Resources/cancel.gif", "Delete this language entry.");
				delete.Margin = new Thickness(0, 0, 8, 0);

				var stack = new StackPanel();
				stack.Orientation = Orientation.Horizontal;
				stack.Children.Add(delete);
				stack.Children.Add(checkbox);

				var defaultDict = new System.Windows.Controls.Primitives.ToggleButton();
				defaultDict.IsThreeState = false;
				if (!lang.HasStandardDictionary) defaultDict.Visibility = System.Windows.Visibility.Hidden;
				defaultDict.IsChecked = lang.CustomDictionaries.Contains("*");
				defaultDict.Content = Image("Resources/anchor.png", "Use the standard dictionary.");

				var customDict = new System.Windows.Controls.Primitives.ToggleButton();
				customDict.IsThreeState = false;
				customDict.IsChecked = lang.CustomDictionaries.Any(d => d != "*");
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
					lang.Enabled = checkbox.IsChecked ?? false;
				};

				delete.Click += (sender, args) => {
					Configuration.Languages.Remove(lang);
					panel.Children.Remove(dock);
				};

				defaultDict.Click += (sender, args) => {
					var has = lang.CustomDictionaries.Any(d => d == "*");
 					if (has != (defaultDict.IsChecked ?? false)) {
						if (!has) lang.CustomDictionaries.Add("*");
						else lang.CustomDictionaries.Remove("*");
					}
				};

				customDict.Click += (sender, args) => {
					if (customDict.IsChecked ?? false) {
						// show file dialog to select custom dictionary files.
						var d = new System.Windows.Forms.OpenFileDialog();
						d.Title = "Select Custom Dictionary Files for " + lang.Culture.EnglishName;
						d.SupportMultiDottedExtensions = true;
						d.Multiselect = true;
						d.Filter = "Dictionary Files (*.lex;*.txt;*.dic)|*.lex;*.txt;*.dic|All files (*.*)|*.*";
						d.ShowDialog();
						foreach (var file in d.FileNames) {
							var newfile = System.IO.Path.Combine(Configuration.ConfigDirectory, System.IO.Path.GetFileName(file));
							var ext = System.IO.Path.GetExtension(file);
							if (ext == ".dic") { // import dic files from NetSpell & ISpell
								newfile = Configuration.ImportDic(file);
							} else {
								System.IO.File.Copy(file, newfile);
							}
							lang.CustomDictionaries.Add(System.IO.Path.GetFileName(newfile));
						}
					} else {
  						foreach (var customd in lang.CustomDictionaries.Where(d => d != "*").ToList()) {
							var file = System.IO.Path.Combine(Configuration.ConfigDirectory, customd);
							if (File.Exists(file)) File.Delete(file);
							lang.CustomDictionaries.Remove(customd);
						}
					}
				};

				panel.Children.Insert(index, dock);
			}
		}

	}
}
