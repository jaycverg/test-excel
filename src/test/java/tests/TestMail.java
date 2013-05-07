package tests;

import java.util.Date;
import java.util.Properties;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class TestMail
{

    @Test
    public void test()
    {
        String to = "jaycverg@gmail.com";
        String from = "jaycverg@gmail.com";

        Properties props = new Properties();
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.socketFactory.port", "465");
        props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.port", "465");
        props.put("mail.debug", "true");
        
        Session session = Session.getInstance(props);

        try {
            // Instantiate a message
            MimeMessage emailMsg = new MimeMessage(session);

            //Set message attributes
            emailMsg.setFrom(new InternetAddress(from));
            InternetAddress[] address = {new InternetAddress(to)};
            emailMsg.setRecipients(Message.RecipientType.TO, address);
            emailMsg.setSubject("Test E-Mail through Java");
            emailMsg.setSentDate(new Date());

            String  htmlMsg = "<html><b>This</b> is the content</html>";
            emailMsg.setContent(htmlMsg, "text/html");

            //Send the message
            Transport transport = session.getTransport("smtp");
            transport.connect("jaycverg@gmail.com", "");
            transport.sendMessage(emailMsg, emailMsg.getRecipients(Message.RecipientType.TO));
            transport.close();
        }
        catch (MessagingException mex) {
            // Prints all nested (chained) exceptions as well
            mex.printStackTrace();
        }
    }
}
