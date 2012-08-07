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
import java.util.logging.Level;
import java.util.logging.Logger;
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
    public static boolean processedFile = false;
    public static boolean movedFile = false;

    private static boolean process(File dir) {
        try {
            PdfReader ReadInputPDF = new PdfReader(dir.getAbsolutePath());
            HashMap<String, String> metaDataInfo = ReadInputPDF.getInfo();
            System.out.println(metaDataInfo.get("Keywords"));
            sendEmail(dir.getAbsolutePath(),metaDataInfo.get("Keywords"));
            return true;
            /* dumping metadata on the screen */
        } catch (Exception ex) {
            Logger.getLogger(EWSConnection.class.getName()).log(Level.SEVERE, "Exception while sending email : ", ex);
            return false;
        } 
        
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
        msg.setSubject("Itemisation Report - ResilientPLC");
        msg.setBody(MessageBody.getMessageBodyFromText("Sent using the EWS Managed API."));
        msg.getToRecipients().add(Email);
        msg.getAttachments().addFileAttachment(path);
        msg.send();
    }

    public static void processFolder(File dir)  {
        if (dir.isDirectory()) {
            String[] children = dir.list();
            File file = null;
            File archiveDir = new File("C:/Archive");
            File errorDir = new File("C:/Error");
            for (int i = 0; i < children.length; i++) {
                file = new File(dir.getAbsoluteFile()+"\\"+children[i]);
                processedFile = process(file);
                if (processedFile){
                    movedFile = file.renameTo(new File(archiveDir, file.getName()));
                }else{
                    movedFile = file.renameTo(new File(errorDir, file.getName()));
                }
            }
        }
    }
}
