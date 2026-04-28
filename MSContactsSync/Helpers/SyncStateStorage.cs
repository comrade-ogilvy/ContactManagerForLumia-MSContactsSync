// Helpers/SyncStateStorage.cs
// Stores per-contact ETag in ApplicationData.LocalSettings

using System;
using System.Collections.Generic;
using Windows.Storage;

namespace MSContactsSync.Helpers
{
    public static class SyncStateStorage
    {
        private static ApplicationDataContainer Settings =>
            ApplicationData.Current.LocalSettings;

        private const string Key = "MsSyncStateV1";

        // Format: id\tetag\n per contact
        public static void Save(Dictionary<string, string> etags)
        {
            try
            {
                var sb = new System.Text.StringBuilder();
                foreach (var kv in etags)
                {
                    if (string.IsNullOrEmpty(kv.Key)) continue;
                    sb.Append(Esc(kv.Key));   sb.Append('\t');
                    sb.Append(Esc(kv.Value)); sb.Append('\n');
                }
                Settings.Values[Key] = sb.ToString();
            }
            catch { }
        }

        public static Dictionary<string, string> Load()
        {
            var result = new Dictionary<string, string>();
            try
            {
                if (!Settings.Values.ContainsKey(Key)) return result;
                string raw = Settings.Values[Key] as string;
                if (string.IsNullOrEmpty(raw)) return result;
                foreach (string line in raw.Split('\n'))
                {
                    if (string.IsNullOrEmpty(line)) continue;
                    int sep = line.IndexOf('\t');
                    if (sep < 0) continue;
                    string id   = Unesc(line.Substring(0, sep));
                    string etag = Unesc(line.Substring(sep + 1));
                    if (!string.IsNullOrEmpty(id))
                        result[id] = etag;
                }
            }
            catch { }
            return result;
        }

        public static void Clear()
        {
            Settings.Values.Remove(Key);
        }

        // DeltaLink for incremental sync
        public static void SaveDeltaLink(string deltaLink)
            => Settings.Values["MsDeltaLink"] = deltaLink ?? "";

        public static string LoadDeltaLink()
            => Settings.Values.ContainsKey("MsDeltaLink")
               ? Settings.Values["MsDeltaLink"] as string : null;

        public static void ClearDeltaLink()
            => Settings.Values.Remove("MsDeltaLink");

        private static string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\")
                    .Replace("\t", "\\t")
                    .Replace("\n", "\\n");
        }

        private static string Unesc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\t",  "\t")
                    .Replace("\\n",  "\n")
                    .Replace("\\\\", "\\");
        }
    }
}
