package phish;

import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.AttachmentCollection;
import microsoft.exchange.webservices.data.property.complex.ItemAttachment;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

public class MailChecker {
  public static void main(String[] args) {
    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
    ExchangeCredentials credentials = new WebCredentials("ITS_phish.svc@ad.unc.edu",
        "&f%n:!jgZ5bk3kux");

    service.setCredentials(credentials);
    try {
      service.autodiscoverUrl("newphish@unc.edu", new RedirectionUrlCallback());
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
              itemAttachment.load();
              System.out.println("Attachment:\n" + itemAttachment.getItem().getBody());
            } else {
              System.out.println("Unknown attachment type");
            }
          }
        } else {
          System.out.println("Message Body:\n" + item.getBody());
        }
      }
    } catch (ServiceLocalException e) {
      e.printStackTrace();
    } catch (Exception e) {
      e.printStackTrace();
    }
    service.close();
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
