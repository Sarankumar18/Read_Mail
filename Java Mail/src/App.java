
//Import Libraries
//import java.io.*;
import java.util.Properties;
import javax.mail.*;
import io.github.cdimascio.dotenv.*;

public class App{
	public static void check(String user, String password) {
		try {
			Properties prop = new Properties();
			prop.setProperty("mail.store.protocol", "imaps");
			Session emailSession = 	Session.getDefaultInstance(prop);
			Store emailStore = emailSession.getStore("imaps");
			emailStore.connect("outlook.office365.com",user,password);
			
			//Opening Folder
			Folder emailFolder = emailStore.getFolder("INBOX");
			emailFolder.open(Folder.READ_ONLY);
			
			Message messages[] = emailFolder.getMessages();
			
			int i = ((messages.length) -1);
			
			Message message = 	messages[i];
			
			System.out.println("Email Number" + (i+1));
			System.out.println("Subject Line" + message.getSubject());
			
			emailFolder.close(true);
			emailStore.close();
		}
		catch(NoSuchProviderException e) {e.printStackTrace();}
		catch(MessagingException e) {e.printStackTrace();}
	}
	
	public static void main(String[] args) {
		Dotenv dotenv = Dotenv.load();
    String username = dotenv.get("EMAIL");
    String password = dotenv.get("PASSWORD");
		check(username, password);
	}
}