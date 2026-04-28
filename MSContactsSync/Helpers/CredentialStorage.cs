// Helpers/CredentialStorage.cs
// Stores OAuth tokens in ApplicationData.LocalSettings

using Windows.Storage;

namespace MSContactsSync.Helpers
{
    public static class CredentialStorage
    {
        private static ApplicationDataContainer Settings =>
            ApplicationData.Current.LocalSettings;

        public static void SaveClientId(string clientId)
            => Settings.Values["ClientId"] = clientId;

        public static string LoadClientId()
            => Settings.Values.ContainsKey("ClientId")
               ? Settings.Values["ClientId"] as string : null;

        public static void SaveToken(string refreshToken)
            => Settings.Values["RefreshToken"] = refreshToken;

        public static string LoadToken()
            => Settings.Values.ContainsKey("RefreshToken")
               ? Settings.Values["RefreshToken"] as string : null;

        public static bool HasToken()
            => !string.IsNullOrEmpty(LoadToken());

        public static void DeleteToken()
        {
            Settings.Values.Remove("RefreshToken");
            Settings.Values.Remove("AccessToken");
        }

        public static void SaveAccessToken(string accessToken)
            => Settings.Values["AccessToken"] = accessToken;

        public static string LoadAccessToken()
            => Settings.Values.ContainsKey("AccessToken")
               ? Settings.Values["AccessToken"] as string : null;

        public static void SaveExpiry(long expiresOn)
            => Settings.Values["TokenExpiry"] = expiresOn.ToString();

        public static long LoadExpiry()
        {
            if (!Settings.Values.ContainsKey("TokenExpiry")) return 0;
            long v; long.TryParse(Settings.Values["TokenExpiry"] as string, out v);
            return v;
        }

        public static void SaveUsername(string username)
            => Settings.Values["Username"] = username ?? "";

        public static string LoadUsername()
            => Settings.Values.ContainsKey("Username")
               ? Settings.Values["Username"] as string : "";
    }
}
