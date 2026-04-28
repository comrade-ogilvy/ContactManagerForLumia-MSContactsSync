// MainPage.xaml.cs
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel.Email;
using Windows.Storage;
using Windows.UI.Core;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using MSContactsSync.Helpers;
using MSContactsSync.Services;

namespace MSContactsSync
{
    public sealed partial class MainPage : Page
    {
        private readonly StringBuilder _log = new StringBuilder();
        private int _pendingLogUpdate = 0;
        private GraphApiService _api;
        private DispatcherTimer _pollTimer;

        public MainPage()
        {
            InitializeComponent();
            Loaded += MainPage_Loaded;
        }

        private async void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                string savedId = CredentialStorage.LoadClientId();
                TxtClientId.Text = savedId ?? "";

                if (CredentialStorage.HasToken())
                    ShowSignedInState(CredentialStorage.LoadUsername());
                else
                    ShowSignedOutState();
            }
            catch
            {
                ShowSignedOutState();
            }

            // Request contacts permission early so it doesn't crash later
            try
            {
                await Windows.ApplicationModel.Contacts.ContactManager
                    .RequestStoreAsync(
                        Windows.ApplicationModel.Contacts
                            .ContactStoreAccessType.AppContactsReadWrite);
            }
            catch { }
        }

        // ================================================================
        // LOAD JSON CONFIG
        // ================================================================
        private async void BtnLoadJson_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var picker = new Windows.Storage.Pickers.FileOpenPicker();
                picker.SuggestedStartLocation =
                    Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary;
                picker.FileTypeFilter.Add(".json");

                var file = await picker.PickSingleFileAsync();
                if (file == null) return;

                string text = await Windows.Storage.FileIO.ReadTextAsync(file);

                // Try MSAL cache format first
                var cache = JsonHelper.ParseMsalCache(text);
                if (cache.IsValid)
                {
                    CredentialStorage.SaveClientId(cache.ClientId);
                    CredentialStorage.SaveToken(cache.RefreshToken);
                    if (!string.IsNullOrEmpty(cache.Username))
                        CredentialStorage.SaveUsername(cache.Username);
                    if (cache.AccessTokenValid)
                    {
                        CredentialStorage.SaveAccessToken(cache.AccessToken);
                        CredentialStorage.SaveExpiry(cache.ExpiresOn);
                    }
                    TxtClientId.Text = cache.ClientId;
                    ShowSignedInState(cache.Username);
                    Log("Loaded MSAL cache: " + file.Name);
                    TxtLoginStatus.Text = "Signed in as: " +
                        (cache.Username ?? cache.ClientId);
                    return;
                }

                // Try simple config.json format
                if (ParseConfig(text))
                {
                    Log("Loaded config: " + file.Name);
                    TxtLoginStatus.Text = "Client ID loaded from: " + file.Name;
                    return;
                }

                TxtLoginStatus.Text = "No client_id found in: " + file.Name;
            }
            catch (Exception ex)
            {
                TxtLoginStatus.Text = "Error: " + ex.Message;
            }
        }

        private async Task LoadFromDocumentsAsync(bool showResult = false)
        {
            // Try to auto-load MSAL cache from LocalSettings backup path
            try
            {
                var localFolder = Windows.Storage.ApplicationData.Current.LocalFolder;
                var file = await localFolder.GetFileAsync(
                    GraphApiService.CacheFileName);
                string text = await Windows.Storage.FileIO.ReadTextAsync(file);
                var cache = JsonHelper.ParseMsalCache(text);
                if (cache.IsValid)
                {
                    CredentialStorage.SaveClientId(cache.ClientId);
                    CredentialStorage.SaveToken(cache.RefreshToken);
                    if (!string.IsNullOrEmpty(cache.Username))
                        CredentialStorage.SaveUsername(cache.Username);
                    if (cache.AccessTokenValid)
                    {
                        CredentialStorage.SaveAccessToken(cache.AccessToken);
                        CredentialStorage.SaveExpiry(cache.ExpiresOn);
                    }
                    TxtClientId.Text = cache.ClientId;
                    ShowSignedInState(cache.Username);
                    Log("Auto-loaded MSAL cache.");
                }
            }
            catch { }
        }

        private bool ParseConfig(string json)
        {
            try
            {
                string clientId = null;
                var obj = Windows.Data.Json.JsonObject.Parse(json);

                if (obj.ContainsKey("client_id"))
                    clientId = JsonHelper.GetString(obj, "client_id");
                else if (obj.ContainsKey("installed"))
                    clientId = JsonHelper.GetString(
                        obj.GetNamedObject("installed"), "client_id");
                else if (obj.ContainsKey("web"))
                    clientId = JsonHelper.GetString(
                        obj.GetNamedObject("web"), "client_id");

                if (!string.IsNullOrEmpty(clientId))
                {
                    TxtClientId.Text = clientId;
                    CredentialStorage.SaveClientId(clientId);
                    return true;
                }
                return false;
            }
            catch { return false; }
        }

        // ================================================================
        // SIGN IN
        // ================================================================
        private async void BtnSignIn_Click(object sender, RoutedEventArgs e)
        {
            string clientId = TxtClientId.Text.Trim();
            if (string.IsNullOrEmpty(clientId))
            {
                TxtLoginStatus.Text = "Please enter Client ID.";
                return;
            }

            CredentialStorage.SaveClientId(clientId);
            _api = new GraphApiService(clientId);

            BtnSignIn.IsEnabled = false;
            TxtLoginStatus.Text = "Requesting device code...";

            bool ok = await _api.StartDeviceFlowAsync();
            if (!ok)
            {
                TxtLoginStatus.Text = "Failed to start device flow. Check Client ID.";
                BtnSignIn.IsEnabled = true;
                return;
            }

            TxtVerificationUrl.Text = _api.VerificationUrl;
            TxtUserCode.Text        = _api.UserCode;
            PanelCode.Visibility    = Visibility.Visible;
            TxtLoginStatus.Text     = "";

            // Start polling
            _pollTimer = new DispatcherTimer();
            _pollTimer.Interval = TimeSpan.FromSeconds(_api.Interval);
            _pollTimer.Tick    += PollTimer_Tick;
            _pollTimer.Start();
        }

        private async void PollTimer_Tick(object sender, object e)
        {
            string result = await _api.PollForTokenAsync();
            if (result == "ok")
            {
                _pollTimer.Stop();
                PanelCode.Visibility = Visibility.Collapsed;

                // Fetch username from /me
                string accessToken = CredentialStorage.LoadAccessToken();
                if (!string.IsNullOrEmpty(accessToken))
                {
                    string username = await _api.GetUsernameAsync(accessToken);
                    if (!string.IsNullOrEmpty(username))
                        CredentialStorage.SaveUsername(username);
                    ShowSignedInState(username);
                }
                else ShowSignedInState();

                Log("Signed in successfully.");
                TxtLoginStatus.Text = "";

                // Save MSAL cache to Documents
                await SaveMsalCacheAsync();
            }
            else if (result == "authorization_pending" ||
                     result == "slow_down")
            {
                // Keep polling
                if (result == "slow_down")
                    _pollTimer.Interval =
                        TimeSpan.FromSeconds(_pollTimer.Interval.TotalSeconds + 5);
            }
            else
            {
                _pollTimer.Stop();
                PanelCode.Visibility = Visibility.Collapsed;
                TxtLoginStatus.Text  = "Auth failed: " + result;
                BtnSignIn.IsEnabled  = true;
            }
        }

        // ================================================================
        // SYNC — Microsoft → Phone
        // ================================================================
        private void BtnSync_Click(object sender, RoutedEventArgs e)
        {
            SetUiBusy(true);
            Log("=== Sync started: " + DateTime.Now.ToString("HH:mm:ss") + " ===");
            MainPivot.SelectedIndex = 2; // Switch to Log tab

            string clientId = CredentialStorage.LoadClientId();
            if (string.IsNullOrEmpty(clientId))
            {
                Log("Error: no Client ID. Go to Sign in tab.");
                SetUiBusy(false);
                return;
            }

            _api = new GraphApiService(clientId);

            Task.Run(async () =>
            {
                try
                {
                    // Refresh token
                    string refreshToken = CredentialStorage.LoadToken();
                    Log("Getting access token...");
                    string accessToken = await _api.GetAccessTokenAsync(refreshToken);
                    if (string.IsNullOrEmpty(accessToken))
                    {
                        Log("Failed to get access token. Please sign in again.");
                        return;
                    }
                    Log("Access token OK.");

                    // Fetch contacts from Graph
                    Log("Fetching contacts from Microsoft...");
                    var fetchResult =
                        await _api.FetchAllContactsAsync(accessToken,
                            msg => Log(msg));
                    var contacts  = fetchResult.Contacts;
                    var deltaLink = fetchResult.DeltaLink;
                    Log("Fetched: " + contacts.Count + " contacts.");

                    if (contacts.Count == 0)
                    {
                        Log("No contacts found.");
                        return;
                    }

                    // Load saved ETags
                    var savedEtags = SyncStateStorage.Load();
                    bool isFirst   = savedEtags.Count == 0;
                    bool isDelta   = !string.IsNullOrEmpty(
                        SyncStateStorage.LoadDeltaLink());
                    Log("isFirst=" + isFirst + " isDelta=" + isDelta +
                        " savedETags=" + savedEtags.Count);

                    // Process contacts
                    var store    = new ContactStoreService();
                    var newEtags = new Dictionary<string, string>(savedEtags);
                    int updated  = 0;
                    int skipped  = 0;
                    int deleted  = 0;
                    int idx      = 0;

                    foreach (var c in contacts)
                    {
                        if (string.IsNullOrEmpty(c.Id)) continue;
                        idx++;
                        if (idx % 20 == 0)
                            Log("Processing " + idx + "/" + contacts.Count + "...");

                        // Delta query marks deleted contacts
                        if (c.IsDeleted)
                        {
                            await store.DeleteContactAsync(c.Id);
                            newEtags.Remove(c.Id);
                            deleted++;
                            continue;
                        }

                        bool changed = !savedEtags.ContainsKey(c.Id) ||
                                       savedEtags[c.Id] != c.ETag;

                        if (isFirst || !isDelta || changed)
                        {
                            await store.UpsertContactAsync(c);
                            newEtags[c.Id] = c.ETag ?? "";
                            updated++;
                        }
                        else skipped++;
                    }

                    // Save state + deltaLink for next incremental sync
                    SyncStateStorage.Save(newEtags);
                    if (!string.IsNullOrEmpty(deltaLink))
                    {
                        SyncStateStorage.SaveDeltaLink(deltaLink);
                        Log("DeltaLink saved for next sync.");
                    }

                    // Update MSAL cache file in Documents
                    await Dispatcher.RunAsync(
                        Windows.UI.Core.CoreDispatcherPriority.Normal,
                        async () => await SaveMsalCacheAsync());

                    Log("Updated=" + updated + " Skipped=" + skipped +
                        " Deleted=" + deleted);
                    Log("=== Done: " + DateTime.Now.ToString("HH:mm:ss") + " ===");

                    await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
                    {
                        TxtLastSync.Text       = "Last sync: " +
                            DateTime.Now.ToString("dd MMM yyyy HH:mm");
                        TxtLastSync.Visibility = Visibility.Visible;
                    });
                }
                catch (Exception ex)
                {
                    Log("EXCEPTION: " + ex.GetType().Name);
                    Log("MSG: " + ex.Message);
                    if (ex.InnerException != null)
                        Log("INNER: " + ex.InnerException.Message);
                }
                finally
                {
                    FlushLog();
                    await Dispatcher.RunAsync(CoreDispatcherPriority.Normal,
                        () => SetUiBusy(false));
                }
            });
        }

        // ================================================================
        // SAVE MSAL CACHE to Documents
        // ================================================================
        private async Task SaveMsalCacheAsync()
        {
            try
            {
                string clientId     = CredentialStorage.LoadClientId();
                string refreshToken = CredentialStorage.LoadToken();
                string accessToken  = CredentialStorage.LoadAccessToken();
                string username     = CredentialStorage.LoadUsername();
                long   expiresOn    = CredentialStorage.LoadExpiry();

                if (string.IsNullOrEmpty(refreshToken)) return;

                string json = _api.BuildMsalCacheJson(
                    clientId, refreshToken, accessToken, username, expiresOn);

                // Save to app's local folder — no special permissions needed
                var localFolder = Windows.Storage.ApplicationData.Current.LocalFolder;
                var file = await localFolder.CreateFileAsync(
                    GraphApiService.CacheFileName,
                    Windows.Storage.CreationCollisionOption.ReplaceExisting);
                await Windows.Storage.FileIO.WriteTextAsync(file, json);
                Log("MSAL cache saved.");
            }
            catch (Exception ex)
            {
                Log("Cache save error: " + ex.Message);
            }
        }

        // ================================================================
        // SIGN OUT
        // ================================================================
        private async void BtnSignOut_Click(object sender, RoutedEventArgs e)
        {
            CredentialStorage.DeleteToken();
            SyncStateStorage.Clear();
            SyncStateStorage.ClearDeltaLink();
            ShowSignedOutState();
            TxtLastSync.Visibility = Visibility.Collapsed;

            // Delete cache from local folder
            try
            {
                var localFolder = Windows.Storage.ApplicationData.Current.LocalFolder;
                var file = await localFolder.GetFileAsync(
                    GraphApiService.CacheFileName);
                await file.DeleteAsync();
            }
            catch { }
        }

        // ================================================================
        // EMAIL LOG
        // ================================================================
        private async void BtnEmailLog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var msg = new EmailMessage();
                msg.Subject = "MSContactSync log " +
                    DateTime.Now.ToString("dd MMM yyyy HH:mm");
                msg.Body    = _log.ToString();
                await EmailManager.ShowComposeNewEmailAsync(msg);
            }
            catch { }
        }

        // ================================================================
        // CLEAR LOG
        // ================================================================
        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            _log.Clear();
            System.Threading.Interlocked.Exchange(ref _pendingLogUpdate, 0);
            TxtLog.Text = "";
        }

        // ================================================================
        // UI STATE
        // ================================================================
        private void ShowSignedInState(string username = null)
        {
            string who = string.IsNullOrEmpty(username)
                ? "Microsoft account connected"
                : "Signed in: " + username;
            TxtAccountStatus.Text = who;
            BtnSync.IsEnabled     = true;
            BtnSignOut.IsEnabled  = true;
            BtnSignIn.IsEnabled   = false;
        }

        private void ShowSignedOutState()
        {
            TxtAccountStatus.Text = "Not signed in";
            BtnSync.IsEnabled     = false;
            BtnSignOut.IsEnabled  = false;
            BtnSignIn.IsEnabled   = true;
        }

        private void SetUiBusy(bool busy)
        {
            BtnSync.IsEnabled    = !busy;
            BtnSignOut.IsEnabled = !busy;
            BtnSignIn.IsEnabled  = !busy;
        }

        // ================================================================
        // LOGGING
        // ================================================================
        private void Log(string msg)
        {
            string line = DateTime.Now.ToString("HH:mm:ss") + " " + msg;
            _log.AppendLine(line);

            if (System.Threading.Interlocked.Exchange(ref _pendingLogUpdate, 1) == 0)
            {
                var _ = Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
                {
                    System.Threading.Interlocked.Exchange(ref _pendingLogUpdate, 0);
                    TxtLog.Text = _log.ToString();
                    LogScroller.ChangeView(null,
                        LogScroller.ScrollableHeight, null);
                });
            }
        }

        private void FlushLog()
        {
            var _ = Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () =>
            {
                TxtLog.Text = _log.ToString();
                LogScroller.ChangeView(null, LogScroller.ScrollableHeight, null);
            });
        }
    }
}
