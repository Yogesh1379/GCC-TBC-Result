package SendEmail_Package;

import java.io.File;
import java.util.Properties;

import jakarta.mail.*;
import jakarta.mail.internet.*;

public class SendMailofSpeedQuestionVerify {

    public static void sendReport(String reportPath)
            throws Exception {

        final String senderEmail ="yogesh.dangade@winnersoft.co.in";

        final String password ="mnkj ddcp nblf ycqd";

        String receiverEmail = "yogeshdangade123@gmail.com";

        Properties props = new Properties();

        props.put("mail.smtp.host", "smtp.gmail.com");

        props.put("mail.smtp.port", "587");

        props.put("mail.smtp.auth", "true");

        props.put("mail.smtp.starttls.enable", "true");

        Session session = Session.getInstance(
                props,
                new Authenticator() {

                    protected PasswordAuthentication
                    getPasswordAuthentication() {

                        return new PasswordAuthentication(
                                senderEmail,
                                password);
                    }
                });

        Message message =
                new MimeMessage(session);

        message.setFrom(
                new InternetAddress(senderEmail));

        message.setRecipients(
                Message.RecipientType.BCC,
                InternetAddress.parse(receiverEmail));

        message.setSubject(
                "Question Paper Verification Report - July 2026");

        // Mail Body
        MimeBodyPart messageBodyPart =
                new MimeBodyPart();

        messageBodyPart.setText(
                "Dear Team,\r\n"
                + "\r\n"
                + "Please find attached the Speed Question Paper Verification Report generated through automation.\r\n"
                + "\r\n"
                + "The report contains formatting and validation issues identified in the question papers.\r\n"
                + "\r\n"
                + "Regards,\r\n"
                + "QA Automation Team");

        // Attachment
        MimeBodyPart attachmentPart =
                new MimeBodyPart();

        attachmentPart.attachFile(
                new File(reportPath));

        Multipart multipart =
                new MimeMultipart();

        multipart.addBodyPart(messageBodyPart);

        multipart.addBodyPart(attachmentPart);

        message.setContent(multipart);

        Transport.send(message);

        System.out.println("Mail Sent Successfully");
    }
}