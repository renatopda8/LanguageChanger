using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Forms;
using Application = System.Windows.Application;
using MessageBox = System.Windows.MessageBox;

namespace LanguageChanger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            CbLanguage.ItemsSource = Languages;
            CbLanguage.DisplayMemberPath = nameof(LanguageChanger.Language.Name);
            CbLanguage.SelectedValuePath = nameof(LanguageChanger.Language.GameTag);

            LanguageChanged();
        }

        private void LanguageChanged()
        {
            Match match = LocaleRegex.Match(Settings);
            if (!match.Success || match.Groups.Count <= 1)
            {
                SettingsLanguage = null;
                return;
            }

            SettingsLanguage = Languages.First(l => l.GameTag.Equals(match.Groups[1].Value));
        }

        private static string FindLeaguePath()
        {
            FolderBrowser.ShowDialog();
            if (File.Exists($"{FolderBrowser.SelectedPath}\\{SettingsFile}"))
            {
                return FolderBrowser.SelectedPath;
            }
            if (MessageBox.Show("Settings file not found. Do you want to select a new folder?", "Error",
                MessageBoxButton.YesNo, MessageBoxImage.Error) == MessageBoxResult.Yes)
            {
                return FindLeaguePath();
            }
            Application.Current.Shutdown();
            throw new Exception("Settings file not found");
        }

        private static FolderBrowserDialog _folderBrowser;

        /// <summary>
        /// Folder browser dialog
        /// </summary>
        private static FolderBrowserDialog FolderBrowser
            => _folderBrowser ?? (_folderBrowser = new FolderBrowserDialog
               {
                   ShowNewFolderButton = false,
                   Description = "League of Legends folder"
               });

        private Language _settingsLanguage;
        /// <summary>
        /// Language in the settings file
        /// </summary>
        private Language SettingsLanguage
        {
            get { return _settingsLanguage; }
            set
            {
                _settingsLanguage = value;
                CbLanguage.SelectedItem = SettingsLanguage;
            }
        }

        private static Regex _localeRegex;
        /// <summary>
        /// Locale regular expression
        /// </summary>
        private static Regex LocaleRegex => _localeRegex ?? (_localeRegex = new Regex("locale: \"(.._..)\""));

        private static string _leaguePath;
        /// <summary>
        /// League of Legends path
        /// </summary>
        private static string LeaguePath => _leaguePath ?? (_leaguePath = FindLeaguePath());        

        /// <summary>
        /// Settings file
        /// </summary>
        private const string SettingsFile = "Config\\LeagueClientSettings.yaml";

        /// <summary>
        /// Setting path
        /// </summary>
        private static string SettingsPath => $"{LeaguePath}\\{SettingsFile}";

        private string _settings;
        /// <summary>
        /// League Client Settings
        /// </summary>
        public string Settings
        {
            get { return _settings ?? (_settings = File.ReadAllText(SettingsPath)); }
            private set
            {
                _settings = value;
                LanguageChanged();
            }
        }

        private ObservableCollection<Language> _languages;

        /// <summary>
        /// Available languages
        /// </summary>
        public ObservableCollection<Language> Languages =>
            _languages ?? (_languages = new ObservableCollection<Language>
            {
                new Language {Name = "English (US)", GameTag = "en_US"},
                new Language {Name = "Português", GameTag = "pt_BR"},
                new Language {Name = "Türkçe", GameTag = "tr_TR"},
                new Language {Name = "English (GB)", GameTag = "en_GB"},
                new Language {Name = "Deutsch", GameTag = "de_DE"},
                new Language {Name = "Español (ES)", GameTag = "es_ES"},
                new Language {Name = "Français", GameTag = "fr_FR"},
                new Language {Name = "Italiano", GameTag = "it_IT"},
                new Language {Name = "Čeština", GameTag = "cs_CZ"},
                new Language {Name = "Ελληνικά", GameTag = "el_GR"},
                new Language {Name = "Magyar", GameTag = "hu_HU"},
                new Language {Name = "Polski", GameTag = "pl_PL"},
                new Language {Name = "Română", GameTag = "ro_RO"},
                new Language {Name = "Русский", GameTag = "ru_RU"},
                new Language {Name = "Español (MX)", GameTag = "es_MX"},
                new Language {Name = "English (AU)", GameTag = "en_AU"},
                new Language {Name = "日本語", GameTag = "ja_JP"}
            });

        private void BtSave_Click(object sender, RoutedEventArgs e)
        {
            string oldLanguage = SettingsLanguage.Name;
            Settings = LocaleRegex.Replace(Settings, $"locale: \"{CbLanguage.SelectedValue}\"");
            File.WriteAllText(SettingsPath, Settings);
            MessageBox.Show( $"Language changed from {oldLanguage} to {SettingsLanguage.Name}.", "Sucess", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}