/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package ewsconnect;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.Iterator;
import microsoft.exchange.webservices.data.ConflictResolutionMode;
import microsoft.exchange.webservices.data.EmailMessage;
import microsoft.exchange.webservices.data.EmailMessageSchema;
import microsoft.exchange.webservices.data.ExchangeCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.FindItemsResults;
import microsoft.exchange.webservices.data.Item;
import microsoft.exchange.webservices.data.ItemId;
import microsoft.exchange.webservices.data.ItemSchema;
import microsoft.exchange.webservices.data.ItemView;
import microsoft.exchange.webservices.data.MessageBody;
import microsoft.exchange.webservices.data.SearchFilter;
import microsoft.exchange.webservices.data.ServiceLocalException;
import microsoft.exchange.webservices.data.SortDirection;
import microsoft.exchange.webservices.data.WebCredentials;
import microsoft.exchange.webservices.data.WellKnownFolderName;

/**
 *
 * @author Nimil
 */
public class EWSConnect {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws URISyntaxException, ServiceLocalException, Exception {
        // TODO code application logic here
        ExchangeService service = new ExchangeService();
        ExchangeCredentials credentials = new WebCredentials("makepositive@resilientplc.com", "########");
        service.setCredentials(credentials);
        URI uri = new URI("https://webmail.resilientplc.com/ews/exchange.asmx");
        service.setUrl(uri);
        ItemView view = new ItemView(100);    //Read maximum of 100 emails
        view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending);    // Read emails by received date order by Ascending
        FindItemsResults< Item> results = service.findItems(WellKnownFolderName.Inbox, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false), view);    //Read only unread emails in inbox
        Iterator<Item> itr = results.iterator();
        System.out.println("Total Unread Emails=" + results.getTotalCount());
        while (itr.hasNext()) {
            Item item = itr.next();
            ItemId itemId = item.getId();
            EmailMessage email = EmailMessage.bind(service, itemId);
            System.out.println("Sender= " + email.getSender());
            System.out.println("Subject= " + email.getSubject());
            System.out.println("Body= " + email.getBody());
            email.setIsRead(true);        //Set the email to read.
            email.update(ConflictResolutionMode.AlwaysOverwrite);
        }
        //Send Email 
        EmailMessage msg = new EmailMessage(service);
        msg.setSubject("Hello world!"); 
        msg.setBody(MessageBody.getMessageBodyFromText("Sent using the EWS Managed API."));
        msg.getToRecipients().add("nimil.christopher@makepositive.com");
        msg.getAttachments().addFileAttachment("C:\\Users\\Nimil\\Documents\\NetBeansProjects\\JavaApplication2\\src\\javaapplication2\\m.pdf");
        msg.send();
    }
}