// Services/ContactStoreService.cs
// Uses ContactList API — the correct way to save/delete in UWP:
//   ContactManager.RequestStoreAsync → ContactStore
//   store.FindContactListsAsync → find or create app's ContactList
//   contactList.SaveContactAsync(contact) — saves
//   contactList.DeleteContactAsync(contact) — deletes
//   store.GetContactByIdAsync(id) — find by store ID
//   Contact.RemoteId — set to our own ID for lookup

using System;
using System.Threading.Tasks;
using Windows.ApplicationModel.Contacts;
using MSContactsSync.Models;

namespace MSContactsSync.Services
{
    public class ContactStoreService
    {
        private ContactStore _store;
        private ContactList  _list;
        private const string ListName      = "MSContactsSync";
        private const string RemoteIdPrefix = "MSContactsSync_";

        private async Task<ContactList> GetListAsync()
        {
            if (_list != null) return _list;

            _store = await ContactManager.RequestStoreAsync(
                ContactStoreAccessType.AppContactsReadWrite);

            // Find existing list or create new
            var lists = await _store.FindContactListsAsync();
            foreach (var l in lists)
                if (l.DisplayName == ListName)
                { _list = l; return _list; }

            _list = await _store.CreateContactListAsync(ListName);
            return _list;
        }

        // ================================================================
        // UPSERT — save or update contact
        // ================================================================
        public async Task UpsertContactAsync(MsContact mc)
        {
            try
            {
                var    list = await GetListAsync();
                string rid  = RemoteIdPrefix + mc.Id;

                // Find existing by RemoteId
                Contact contact = null;
                try
                {
                    var reader = list.GetContactReader();
                    while (true)
                    {
                        var batch = await reader.ReadBatchAsync();
                        if (batch.Contacts.Count == 0) break;
                        foreach (var c in batch.Contacts)
                            if (c.RemoteId == rid)
                            { contact = c; break; }
                        if (contact != null) break;
                    }
                }
                catch { }

                if (contact == null) contact = new Contact();
                contact.RemoteId = rid;

                // Name
                contact.FirstName  = mc.FirstName  ?? "";
                contact.MiddleName = mc.MiddleName ?? "";
                contact.LastName   = mc.LastName   ?? "";

                if (string.IsNullOrEmpty(contact.FirstName) &&
                    string.IsNullOrEmpty(contact.LastName)  &&
                    !string.IsNullOrEmpty(mc.DisplayName))
                    contact.LastName = mc.DisplayName;

                // Job
                contact.JobInfo.Clear();
                if (!string.IsNullOrEmpty(mc.Company) ||
                    !string.IsNullOrEmpty(mc.JobTitle))
                    contact.JobInfo.Add(new ContactJobInfo
                    {
                        CompanyName = mc.Company  ?? "",
                        Title       = mc.JobTitle ?? ""
                    });

                // Phones
                contact.Phones.Clear();
                if (!string.IsNullOrEmpty(mc.MobilePhone))
                    contact.Phones.Add(new ContactPhone
                        { Number = mc.MobilePhone, Kind = ContactPhoneKind.Mobile });
                foreach (var p in mc.BusinessPhones)
                    if (!string.IsNullOrEmpty(p))
                        contact.Phones.Add(new ContactPhone
                            { Number = p, Kind = ContactPhoneKind.Work });
                foreach (var p in mc.HomePhones)
                    if (!string.IsNullOrEmpty(p))
                        contact.Phones.Add(new ContactPhone
                            { Number = p, Kind = ContactPhoneKind.Home });

                // Emails
                contact.Emails.Clear();
                foreach (var e in mc.Emails)
                    if (!string.IsNullOrEmpty(e.Address))
                        contact.Emails.Add(new ContactEmail
                            { Address = e.Address, Kind = ContactEmailKind.Personal });

                // Addresses
                contact.Addresses.Clear();
                foreach (var a in mc.Addresses)
                {
                    var kind = ContactAddressKind.Other;
                    if (a.Type == "work") kind = ContactAddressKind.Work;
                    if (a.Type == "home") kind = ContactAddressKind.Home;
                    contact.Addresses.Add(new ContactAddress
                    {
                        Kind          = kind,
                        StreetAddress = a.Street          ?? "",
                        Locality      = a.City            ?? "",
                        Region        = a.State           ?? "",
                        PostalCode    = a.PostalCode       ?? "",
                        Country       = a.CountryOrRegion  ?? ""
                    });
                }

                await list.SaveContactAsync(contact);
            }
            catch { }
        }

        // ================================================================
        // DELETE contact by Graph ID
        // ================================================================
        public async Task DeleteContactAsync(string graphId)
        {
            try
            {
                var    list = await GetListAsync();
                string rid  = RemoteIdPrefix + graphId;

                var reader = list.GetContactReader();
                while (true)
                {
                    var batch = await reader.ReadBatchAsync();
                    if (batch.Contacts.Count == 0) break;
                    foreach (var c in batch.Contacts)
                        if (c.RemoteId == rid)
                        {
                            await list.DeleteContactAsync(c);
                            return;
                        }
                }
            }
            catch { }
        }

        // ================================================================
        // DELETE all app contacts
        // ================================================================
        public async Task DeleteAllContactsAsync(Action<string> progress = null)
        {
            try
            {
                var list    = await GetListAsync();
                var reader  = list.GetContactReader();
                int deleted = 0;
                while (true)
                {
                    var batch = await reader.ReadBatchAsync();
                    if (batch.Contacts.Count == 0) break;
                    foreach (var c in batch.Contacts)
                    {
                        await list.DeleteContactAsync(c);
                        deleted++;
                        if (progress != null && deleted % 10 == 0)
                            progress("Deleted " + deleted + "...");
                    }
                }
            }
            catch { }
        }
    }
}
