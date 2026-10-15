using System.Globalization;

namespace OfficeCliAgent.Core
{
    /// <summary>
    /// Language of the add-in's own UI (ribbon, task pane, status text): Japanese when Office
    /// runs in Japanese, English otherwise. The add-in calls <see cref="Initialize"/> with the
    /// Office UI language; until then the thread's UI culture is used.
    /// </summary>
    public static class UiLanguage
    {
        private static CultureInfo _culture = CultureInfo.CurrentUICulture;

        /// <param name="officeUiLcid">Office UI language (LanguageSettings.LanguageID[msoLanguageIDUI]).</param>
        public static void Initialize(int officeUiLcid)
        {
            try { _culture = CultureInfo.GetCultureInfo(officeUiLcid); }
            catch (CultureNotFoundException) { /* keep the current culture */ }
        }

        public static bool IsJapanese => _culture.TwoLetterISOLanguageName == "ja";

        /// <summary>"ja" or "en"; passed to the chat UI as <c>?lang=</c>.</summary>
        public static string Code => IsJapanese ? "ja" : "en";

        /// <summary>Pick the Japanese or English variant of a UI string.</summary>
        public static string T(string ja, string en) => IsJapanese ? ja : en;
    }
}
