// Services/JsonHelper.cs
using System.Collections.Generic;
using Windows.Data.Json;
using MSContactsSync.Models;

namespace MSContactsSync.Services
{
    // Simple class replacing (string Id, string Name) tuple
    public class FolderInfo
    {
        public string Id   { get; set; }
        public string Name { get; set; }
    }

    public class MsalCacheData
    {
        public string ClientId     { get; set; }
        public string RefreshToken { get; set; }
        public string AccessToken  { get; set; }
        public long   ExpiresOn    { get; set; }
        public string Username     { get; set; }

        public bool IsValid =>
            !string.IsNullOrEmpty(RefreshToken) &&
            !string.IsNullOrEmpty(ClientId);

        public bool AccessTokenValid
        {
            get
            {
                if (string.IsNullOrEmpty(AccessToken)) return false;
                return ExpiresOn > System.DateTimeOffset.UtcNow.ToUnixTimeSeconds() + 60;
            }
        }
    }

    public static class JsonHelper
    {
        // ================================================================
        // PARSE contacts from /me/contacts response
        // ================================================================
        public static List<MsContact> ParseContacts(string json)
        {
            var list = new List<MsContact>();
            try
            {
                var obj   = JsonObject.Parse(json);
                var value = obj.GetNamedArray("value", new JsonArray());
                foreach (var item in value)
                {
                    var c = ParseContact(item.GetObject());
                    if (c != null) list.Add(c);
                }
            }
            catch { }
            return list;
        }

        public static MsContact ParseContact(JsonObject obj)
        {
            try
            {
                var c = new MsContact();
                c.Id          = GetString(obj, "id");
                c.ETag        = GetString(obj, "@odata.etag");
                // Delta query marks deleted contacts with @removed
                c.IsDeleted   = obj.ContainsKey("@removed");
                c.DisplayName = GetString(obj, "displayName");
                c.FirstName   = GetString(obj, "givenName");
                c.MiddleName  = GetString(obj, "middleName");
                c.LastName    = GetString(obj, "surname");
                c.Nickname    = GetString(obj, "nickName");
                c.Company     = GetString(obj, "companyName");
                c.Department  = GetString(obj, "department");
                c.JobTitle    = GetString(obj, "jobTitle");
                c.Notes       = GetString(obj, "personalNotes");
                c.MobilePhone = GetString(obj, "mobilePhone");

                string bday = GetString(obj, "birthday");
                if (!string.IsNullOrEmpty(bday) && bday.Contains("T"))
                    bday = bday.Split('T')[0];
                c.Birthday = bday;

                if (obj.ContainsKey("businessPhones"))
                    foreach (var p in obj.GetNamedArray("businessPhones"))
                        if (p.ValueType == JsonValueType.String)
                            c.BusinessPhones.Add(p.GetString());

                if (obj.ContainsKey("homePhones"))
                    foreach (var p in obj.GetNamedArray("homePhones"))
                        if (p.ValueType == JsonValueType.String)
                            c.HomePhones.Add(p.GetString());

                if (obj.ContainsKey("emailAddresses"))
                    foreach (var e in obj.GetNamedArray("emailAddresses"))
                    {
                        var eo   = e.GetObject();
                        string a = GetString(eo, "address");
                        if (!string.IsNullOrEmpty(a))
                            c.Emails.Add(new MsEmail
                            {
                                Name    = GetString(eo, "name"),
                                Address = a
                            });
                    }

                ParseAddress(obj, "businessAddress", "work",  c.Addresses);
                ParseAddress(obj, "homeAddress",     "home",  c.Addresses);
                ParseAddress(obj, "otherAddress",    "other", c.Addresses);

                if (string.IsNullOrEmpty(c.Id)) return null;
                return c;
            }
            catch { return null; }
        }

        private static void ParseAddress(JsonObject obj, string key,
            string type, List<MsAddress> list)
        {
            if (!obj.ContainsKey(key)) return;
            var ao = obj.GetNamedObject(key, null);
            if (ao == null) return;
            string street  = GetString(ao, "street");
            string city    = GetString(ao, "city");
            string state   = GetString(ao, "state");
            string postal  = GetString(ao, "postalCode");
            string country = GetString(ao, "countryOrRegion");
            if (string.IsNullOrEmpty(street) && string.IsNullOrEmpty(city) &&
                string.IsNullOrEmpty(postal)) return;
            list.Add(new MsAddress
            {
                Type            = type,
                Street          = street,
                City            = city,
                State           = state,
                PostalCode      = postal,
                CountryOrRegion = country
            });
        }

        // ================================================================
        // PARSE paging links
        // ================================================================
        public static string ParseNextLink(string json)
        {
            try
            {
                return GetString(JsonObject.Parse(json), "@odata.nextLink");
            }
            catch { return null; }
        }

        public static string ParseDeltaLink(string json)
        {
            try
            {
                return GetString(JsonObject.Parse(json), "@odata.deltaLink");
            }
            catch { return null; }
        }

        // ================================================================
        // PARSE contact folders — returns list of FolderInfo
        // ================================================================
        public static List<FolderInfo> ParseFolders(string json)
        {
            var list = new List<FolderInfo>();
            try
            {
                var obj   = JsonObject.Parse(json);
                var value = obj.GetNamedArray("value", new JsonArray());
                foreach (var item in value)
                {
                    var fo   = item.GetObject();
                    string id   = GetString(fo, "id");
                    string name = GetString(fo, "displayName");
                    if (!string.IsNullOrEmpty(id))
                        list.Add(new FolderInfo { Id = id, Name = name });
                }
            }
            catch { }
            return list;
        }

        // ================================================================
        // PARSE token response value
        // ================================================================
        public static string ParseTokenValue(string json, string key)
        {
            try
            {
                return GetString(JsonObject.Parse(json), key);
            }
            catch { return null; }
        }

        // ================================================================
        // PARSE MSAL cache file
        // ================================================================
        public static MsalCacheData ParseMsalCache(string json)
        {
            var result = new MsalCacheData();
            try
            {
                var root = JsonObject.Parse(json);

                // client_id from AppMetadata
                if (root.ContainsKey("AppMetadata"))
                {
                    var meta = root.GetNamedObject("AppMetadata");
                    foreach (var key in meta.Keys)
                    {
                        var entry = meta.GetNamedObject(key, null);
                        if (entry == null) continue;
                        string cid = GetString(entry, "client_id");
                        if (!string.IsNullOrEmpty(cid))
                        {
                            result.ClientId = cid;
                            break;
                        }
                    }
                }

                // refresh token
                if (root.ContainsKey("RefreshToken"))
                {
                    var rt = root.GetNamedObject("RefreshToken");
                    foreach (var key in rt.Keys)
                    {
                        var entry = rt.GetNamedObject(key, null);
                        if (entry == null) continue;
                        string secret = GetString(entry, "secret");
                        if (!string.IsNullOrEmpty(secret))
                        {
                            result.RefreshToken = secret;
                            break;
                        }
                    }
                }

                // access token + expiry
                if (root.ContainsKey("AccessToken"))
                {
                    var at = root.GetNamedObject("AccessToken");
                    foreach (var key in at.Keys)
                    {
                        var entry = at.GetNamedObject(key, null);
                        if (entry == null) continue;
                        string secret = GetString(entry, "secret");
                        string exp    = GetString(entry, "expires_on");
                        if (!string.IsNullOrEmpty(secret))
                        {
                            result.AccessToken = secret;
                            long expiresOn;
                            if (long.TryParse(exp, out expiresOn))
                                result.ExpiresOn = expiresOn;
                            break;
                        }
                    }
                }

                // username
                if (root.ContainsKey("Account"))
                {
                    var accs = root.GetNamedObject("Account");
                    foreach (var key in accs.Keys)
                    {
                        var entry = accs.GetNamedObject(key, null);
                        if (entry == null) continue;
                        string uname = GetString(entry, "username");
                        if (!string.IsNullOrEmpty(uname))
                        {
                            result.Username = uname;
                            break;
                        }
                    }
                }
            }
            catch { }
            return result;
        }

        // ================================================================
        // HELPERS
        // ================================================================
        public static string GetString(JsonObject obj, string key)
        {
            if (obj == null || !obj.ContainsKey(key)) return "";
            var v = obj.GetNamedValue(key);
            if (v.ValueType == JsonValueType.String) return v.GetString();
            if (v.ValueType == JsonValueType.Null)   return "";
            return v.Stringify();
        }
    }
}
