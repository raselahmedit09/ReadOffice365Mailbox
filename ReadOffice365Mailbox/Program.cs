using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;

namespace ReadOffice365Mailbox
{
    class Program
    {

        static void Main(string[] args)
        {

           ExchangeService _exchangeService = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
           string _exchangeDomainUser = ConfigurationManager.AppSettings["exchangeDomainUser"];
           string _exchangeDomainUserPassword = ConfigurationManager.AppSettings["exchangeDomainUserPassword"];
           string _exchangeUrl = ConfigurationManager.AppSettings["exchangeUrl"];
           string _exchangeMailSubject = ConfigurationManager.AppSettings["exchangeMailSubject"];

            try
            {
               
                _exchangeService = new ExchangeService
                {
                    Credentials = new WebCredentials(_exchangeDomainUser, _exchangeDomainUserPassword)
                };

                // This is the office365 webservice URL
                _exchangeService.Url = new Uri(_exchangeUrl);

                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                // create archive folder
                FolderId folderId = CreateArchiveFolder(_exchangeService);

                // file download and move email to archive folder 
                List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
                searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, _exchangeMailSubject));
                SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchFilterCollection.ToArray());
                ItemView view = new ItemView(50);
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived);
                view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
                view.Traversal = ItemTraversal.Shallow;
                FindItemsResults<Item> findResults = _exchangeService.FindItems(WellKnownFolderName.Inbox, searchFilter, view);

                if (findResults != null && findResults.Items != null && findResults.Items.Count > 0)
                {
                    foreach (EmailMessage item in findResults)
                    {
                        EmailMessage message = EmailMessage.Bind(_exchangeService, item.Id, new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments, ItemSchema.HasAttachments));
                        foreach (Attachment attachment in message.Attachments)
                          {
                            if (attachment is FileAttachment)
                            {
                                FileAttachment fileAttachment = attachment as FileAttachment;
                                fileAttachment.Load(@"D:\\Files\\" + fileAttachment.Name);
                            }
                        }

                        // move mail to archive folder 
                        item.Move(folderId);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("new ExchangeService failed");
                return;
            }
        }

        private static FolderId  CreateArchiveFolder(ExchangeService _exchangeService)
        {
            FolderId rootArchiveFoldeId = null;
            FolderId subFolderId = null;

            FindFoldersResults archiveFolder = _exchangeService.FindFolders(WellKnownFolderName.Inbox, new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Archive"), new FolderView(1));
            if (archiveFolder.Folders.Count == 0)
            {
                Folder folder = new Folder(_exchangeService);
                folder.DisplayName = "Archive";
                folder.Save(WellKnownFolderName.Inbox);
                rootArchiveFoldeId = folder.Id;
            }
            else
            {
                rootArchiveFoldeId = archiveFolder.Folders[0].Id;
            }

            string subFolderName = DateTime.Now.ToString("MMMM") + " " + DateTime.Now.ToString("yyyy");
            var archiveSubFolderList = _exchangeService.FindFolders(rootArchiveFoldeId, new SearchFilter.IsEqualTo(FolderSchema.DisplayName, subFolderName), new FolderView(1));
            if (archiveSubFolderList.Folders.Count == 0)
            {
                Folder folder = new Folder(_exchangeService);
                folder.DisplayName = subFolderName;
                folder.Save(rootArchiveFoldeId);
                subFolderId = folder.Id;
            }
            else
            {
                subFolderId = archiveSubFolderList.Folders[0].Id;
            }

            return subFolderId;
        }
    }
}
