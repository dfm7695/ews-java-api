package phish;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import utility.Messenger;
import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.DeleteMode;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.AttachmentCollection;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.ItemAttachment;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;

public class MailChecker {
  private static ExchangeService service;
  private static FolderId archive;

  public static void main(String[] args) {
    service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
    ExchangeCredentials credentials = new WebCredentials("ITS_phish.svc@ad.unc.edu",
        "&f%n:!jgZ5bk3kux");

    service.setCredentials(credentials);
    try {
      service.autodiscoverUrl("newphish@unc.edu", new RedirectionUrlCallback());
      archive = getArchive();
      ItemView view = new ItemView(100);
      FindItemsResults<Item> findResults;
      findResults = service.findItems(WellKnownFolderName.Inbox, view);

      for (Item item : findResults.getItems()) {
        item.load();
        if (item.getHasAttachments()) {
          AttachmentCollection attachments = item.getAttachments();

          for (Attachment attachment : attachments) {
            System.out.println("Attachment name:  " + attachment.getName());
            System.out.println(attachment.getContentType());
            if ("message/rfc822".equals(attachment.getContentType())) {
              ItemAttachment itemAttachment = (ItemAttachment) attachment;
              itemAttachment.load(EmailMessageSchema.MimeContent);
              Item attachmentItem = itemAttachment.getItem();
              item.load();
              String subject = attachmentItem.getSubject();
              String filepath = "C:/Users/dfm7695/Downloads" + subject.toLowerCase() + ".eml";
              File file= new File(filepath);
              OutputStream output = new FileOutputStream(file);
              output.write(attachmentItem.getMimeContent().getContent());
              output.close();
              String body = attachmentItem.getBody().toString();
              System.out.println("Attachment:\n" + body);
              System.out.println("Subject:  " + subject);
              System.out.println("Link:  " + getLinks(body).toString());
            } else {
              System.out.println("Unknown attachment type");
            }
          }
        } else {
          String body = item.getBody().toString();
          System.out.println("Message Body:\n" + item.getBody());
          System.out.println("Link:  " + getLinks(body).toString());
        }
        item.copy(archive);
        item.delete(DeleteMode.HardDelete);
      }
    } catch (ServiceLocalException e) {
      e.printStackTrace();
      new Messenger().email("phish.MailChecker.main(ServiceLocalException):\n" + e.getMessage());
    } catch (Exception e) {
      e.printStackTrace();
      new Messenger().email("phish.MailChecker.main(Exception):\n" + e.getMessage());
    }
    service.close();
  }

  private static FolderId getArchive() {
    try {
      Folder rootfolder = Folder.bind(service, WellKnownFolderName.MsgFolderRoot);

      for (Folder folder : rootfolder.findFolders(new FolderView(100))) {
        if (folder.getDisplayName().equals("Archive")) {
          return folder.getId();
        }
      }
    } catch (Exception e) {
      e.printStackTrace();
      new Messenger().email("phish.MailChecker.getArchive(Exception):\n" + e.getMessage());
    }
    return null;
  }

  private static Set<String> getLinks(String body) {
    Set<String> links = new HashSet<String>();
    String regex = "\\(?\\b(https?://|www[.])"
        + "[-A-Za-z0-9+&amp;@#/%?=~_()|!:,.;]*[-A-Za-z0-9+&amp;@#/%=~_()|]";
    Pattern p = Pattern.compile(regex);
    int start = body.indexOf("<body") + 6;
    System.out.println(start);
    String link = body.substring(start);
    Matcher m = p.matcher(link);

    while (m.find()) {
      String urlStr = m.group();
      if (urlStr.startsWith("(") && urlStr.endsWith(")")) {
        urlStr = urlStr.substring(1, urlStr.length() - 1);
      }
      links.add(urlStr);
    }
    return links;
  }

  /**
   * I found this under "Responding to Autodiscover Redirecting" on this page:
   * https://github.com/OfficeDev/ews-java-api/wiki/Getting-Started-Guide
   */
  static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
    public boolean autodiscoverRedirectionUrlValidationCallback(String redirectionUrl) {
      return redirectionUrl.toLowerCase().startsWith("https://");
    }
  }
}
