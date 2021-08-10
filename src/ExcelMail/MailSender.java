package ExcelMail;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class MailSender {

	public static void send() throws IOException {

		// create username and password
		final String fromEmail = "syedfa143@gmail.com";
		final String password = "Abcd@123_1";

		// adding multiple recipients
		ArrayList<String> mails = new ArrayList<>();
		mails.add("aslam1qqqq@gmail.com");
		mails.add("faheeminfo7547@gmail.com");
		mails.add("ibrahimsb03@gmail.com");
		mails.add("samchodankar1128@gmail.com");

		// adding properties
		Properties props = new Properties();

		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.host", "smtp.gmail.com");
		props.put("mail.smtp.port", "587");
		props.put("MY_EMAIL", "syedfa143@gmail.com");

		// Step1: to get the session object..
		Session session = Session.getInstance(props, new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(fromEmail, password);
			}
		});
		MimeMessage message = new MimeMessage(session);
		for (int i = 0; i < mails.size(); i++) {

			try {
				message.setFrom(new InternetAddress(props.getProperty("MY_EMAIL")));
				message.setRecipient(Message.RecipientType.TO, new InternetAddress(mails.get(i)));
				message.setSubject("HIT Batch-2021 || SAMIKSHA TEAM || COVID VACCINE SURVEY");

				// email 		   --> 		name
				// fahim@gmail.com --> 		fahim
				String txt1 = "Hi Sir/Madam " + mails.get(i).split("@")[0] + ","
						+ "\nKindly find the Survey Details of the District: "
						+ "\nChennai Zone : IV (Tondiarpet)"
						+ "\nCity: kodungaiyur"
						+ "\nPincode : 600118" 
						+ "\nSurvey Person : AADHIL"
						+ "\n\n\nThanks and Regards..." 
						+ "\nSyed Faheem B,"
						+ "'nCEO of Health Ministry," 
						+ "\n+91 9940201750";

				Multipart multipart = new MimeMultipart();

				// set text
				MimeBodyPart text1 = new MimeBodyPart();
				text1.setText(txt1);

				// set attachment
				MimeBodyPart file = new MimeBodyPart();
				file.attachFile(Main.highlightedFile);

				// adding body parts to multipart
				multipart.addBodyPart(text1);
				multipart.addBodyPart(file);

				// adding multipart to the message
				message.setContent(multipart);

				// sending the message that we just created
				Transport.send(message);

				System.out.println("Message sent successfully to " + mails.get(i).split("@")[0]);

			} catch (MessagingException e) {

				e.printStackTrace();
			}
		}
	}
}