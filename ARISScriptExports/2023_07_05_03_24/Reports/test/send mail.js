

 pxsSendMail( ["agudelj90@gmail.com"], "ac", "abc" )


///////////////////////////////////////////////////////////////////////////////////////////////////////
// pxsSendMail function
//  parameters
//    recipientsTo (String[]): array of mail addresses that will be put in the To
//    subject (String): subject of the mail
//    mailBody (String): content of the mail
//
// returns (String) : OK or NOK
// 22/07/2021 apply html/text message format https://jiraprd-proximuscorp.msappproxy.net/browse/J185400100-191
///////////////////////////////////////////////////////////////////////////////////////////////////////
    
function pxsSendMail( recipientsTo, subject, mailBody ) {
    
    if ( subject == null || mailBody == null || recipientsTo == null || recipientsTo.length == 0 ) {
        return "NOK";
    }
    
    var recipientAddressesTo = [];
    for ( iRecipient=0; iRecipient < recipientsTo.length; iRecipient++ ) {
        if ( recipientsTo[ iRecipient ] != "" ) {
            recipientAddressesTo.push( new javax.mail.internet.InternetAddress( recipientsTo[ iRecipient ] ) );
        }
    }
    
    try {
       var props = new java.util.Properties();
       var host = "smtp.freesmtpservers.com";
       props.put("mail.smtp.starttls.enable", "false");
       props.put("mail.smtp.host", host);
       //props.put("mail.smtp.user", "agudelj90@gmail.com");
      // props.put("mail.smtp.password", "");
       props.put("mail.smtp.port", "25");
       //props.put("mail.smtp.auth", "true");
    
       //var pass = new javax.mail.PasswordAuthentication("agudelj90@gmail.com", "KoStElA7")   
        
        
        
      var session=javax.mail.Session.getInstance(props, null); 
       
          //var smtpTransport = session.getTransport("smtp");
     
        //smtpTransport.connect(host, "agudelj90@gmail.com", "KoStElA7");
       
       
       var message = new javax.mail.internet.MimeMessage(session);
       message.setFrom(new javax.mail.internet.InternetAddress("ARIS.noreply@test.com")); 
       message.addRecipients(javax.mail.Message.RecipientType.TO, recipientAddressesTo);
       message.setSubject(subject);
       message.setContent(mailBody, "text/html");       
//       message.setText(mailBody); 
        
       javax.mail.Transport.send(message);
       return "OK";
    } 
    catch(e) {
       return "NOK";
    }
}
