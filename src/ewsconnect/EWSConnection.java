package ewsconnect;


import com.itextpdf.text.pdf.PdfReader;
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.HashMap;
import java.util.logging.Level;
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

    
    public EWSConnection() throws URISyntaxException {
        EWSConnection.service = new ExchangeService();
        ExchangeCredentials credentials = new WebCredentials("makepositive@resilientplc.com", "P0s1t1v3");
        service.setCredentials(credentials);
        URI uri = new URI("https://webmail.resilientplc.com/ews/exchange.asmx");
        service.setUrl(uri);
    }

    private boolean process(File file) {
        boolean processed = false;
        try {
            PdfReader ReadInputPDF = new PdfReader(file.getAbsolutePath());
            HashMap<String, String> metaDataInfo = ReadInputPDF.getInfo();
            //System.out.println(metaDataInfo.get("Keywords"));
            processed = emailReport(file.getAbsolutePath(),metaDataInfo.get("Keywords"));
            /* dumping metadata on the screen */
            return processed;
        } catch (Exception ex) {
            java.util.logging.Logger.getLogger(EWSConnection.class.getName()).log(Level.SEVERE, null, ex);
            return processed;
        } 
        
    }
 

    public boolean emailReport(String path,String emailToAddress) throws Exception {
        //Send Email 
        boolean sent = false;
        try{
            EmailMessage msg = new EmailMessage(service);
            msg.setSubject("Itemisation Report - ResilientPLC");

            msg.setBody(MessageBody.getMessageBodyFromText("Sent using the EWS Managed API."));
            msg.getToRecipients().add(emailToAddress);
            msg.getAttachments().addFileAttachment(path);
            msg.send();
            sent = true;
        }catch(Exception e){
            throw new Exception(e);
        }
        finally{
            return sent;
        }
    }
    
    public void sendEmail(String subject, String emailBody ) throws Exception {
        //Send Email 
        
        try{
            EmailMessage msg = new EmailMessage(service);
            msg.setSubject(subject);
            msg.setBody(MessageBody.getMessageBodyFromText(emailBody));
            /*if(appConfig.getErrorSuccessEmailAddress().indexOf(";") > -1){
                String[] emailToAddresses = appConfig.getErrorSuccessEmailAddress().split(";");
                for(String emailAdd : emailToAddresses){
                    msg.getToRecipients().add(emailAdd);
                }
            }else{
                msg.getToRecipients().add(appConfig.getErrorSuccessEmailAddress());
            }*/
            msg.send();
        }catch(Exception e){
            //LOGGER.error("Exception while sending Success/Error email. Cause :" + e.getMessage());
            throw new Exception(e);
        }
    }

    public void processFolder(File dir) throws IOException, Exception {
        
        //LOGGER.info("Entered into Method : processFolder ");
        //LOGGER.info("Total reports in the Reports Directory := " + dir.list().length);
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