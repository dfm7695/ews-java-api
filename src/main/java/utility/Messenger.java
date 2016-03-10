package utility;

import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

public class Messenger {
  private final String RECIPIENT = "david_mason@unc.edu";

  public void email(String body) {
    Properties props = new Properties();
    Session session = Session.getDefaultInstance(props);

    props.put("mail.smtp.host", "relay.unc.edu");
    try {
      InternetAddress recipient = new InternetAddress(RECIPIENT);
      MimeMessage message = new MimeMessage(session);

      message.setFrom(new InternetAddress("its-cybermation@groups.unc.edu"));
      message.addRecipient(Message.RecipientType.TO, recipient);
      message.setSubject("phish SQL error");
      message.setText(body);
      Transport.send(message);
    } catch (MessagingException mex) {
      mex.printStackTrace();
    }
  }
}
