/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package ewsconnect;

import com.itextpdf.text.pdf.PdfReader;
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.HashMap;
import microsoft.exchange.webservices.data.EmailMessage;
import microsoft.exchange.webservices.data.ExchangeCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.MessageBody;
import microsoft.exchange.webservices.data.WebCredentials;

/**
 *
 * @author Nimil
 */
public class EWSConnection {

    public static ExchangeService service = null;

    private static void process(File dir) throws IOException, Exception {
        PdfReader ReadInputPDF = new PdfReader(dir.getAbsolutePath());
        HashMap<String, String> metaDataInfo = ReadInputPDF.getInfo();
        System.out.println(metaDataInfo.get("Keywords"));
        sendEmail(dir.getAbsolutePath(),metaDataInfo.get("Keywords"));
        /* dumping metadata on the screen */
        
    }


    public EWSConnection() throws URISyntaxException {
        EWSConnection.service = new ExchangeService();
        ExchangeCredentials credentials = new WebCredentials("makepositive@resilientplc.com", "P0s1t1v3");
        service.setCredentials(credentials);
        URI uri = new URI("https://webmail.resilientplc.com/ews/exchange.asmx");
        service.setUrl(uri);
    }

    public static void sendEmail(String path,String Email) throws Exception {
        //Send Email 
        EmailMessage msg = new EmailMessage(service);
        msg.setSubject("Hello world!");
        msg.setBody(MessageBody.getMessageBodyFromText("Sent using the EWS Managed API."));
        msg.getToRecipients().add(Email);
        msg.getAttachments().addFileAttachment(path);
        msg.send();
    }

    public static void processFolder(File dir) throws IOException, Exception {
        if (dir.isDirectory()) {
            String[] children = dir.list();
            for (int i = 0; i < children.length; i++) {
                System.out.println(dir.getAbsoluteFile()+"\\"+children[i]);
                process(new File(dir.getAbsoluteFile()+"\\"+children[i]));
            }
        }
    }
}
