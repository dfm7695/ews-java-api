package phish;

import java.util.Date;
import java.util.Properties;

import javax.mail.Flags;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Store;

public class ArchiveCleaner {
  public static void main(String[] args) {
    Properties props = new Properties();
    Session session = Session.getDefaultInstance(props);
    long old = new Date().getTime() - 2592000000L; // 30 days ago

    props.setProperty("mail.store.protocol", "imap");
    try {
      Store store = session.getStore();
      store.connect("outlook.unc.edu", "ITS_phish.svc", "&f%n:!jgZ5bk3kux");

      Folder archive = store.getFolder("Archive");
      archive.open(Folder.READ_WRITE);

      Message[] messages = archive.getMessages();

      for (Message message : messages) {
        if (message.getReceivedDate().getTime() < old) {
          message.setFlag(Flags.Flag.DELETED, true);
        }
      }
      archive.close(true);
    } catch (MessagingException e) {
      e.printStackTrace();
    }
  }
}
