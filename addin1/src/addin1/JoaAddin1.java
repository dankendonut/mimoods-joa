package addin1;

import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Random;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import com.wilutions.com.BackgTask;
import com.wilutions.com.CoClass;
import com.wilutions.com.Dispatch;
import com.wilutions.com.IDispatch;
import com.wilutions.joa.DeclAddin;
import com.wilutions.joa.IconManager;
import com.wilutions.joa.LoadBehavior;
import com.wilutions.joa.OfficeApplication;
import com.wilutions.joa.fx.MessageBox;
import com.wilutions.joa.outlook.OutlookAddin;
import com.wilutions.mslib.office.IRibbonControl;
import com.wilutions.mslib.outlook.Accounts;
import com.wilutions.mslib.outlook.MailItem;
import com.wilutions.mslib.outlook._Account;
import com.wilutions.mslib.outlook._Inspector;

@CoClass(progId = "JoaAddin1.Class", guid = "{d5f0439b-27ea-4848-a230-3fa5496ea5e1}")
@DeclAddin(friendlyName = "My First JOA Add-in", description = "Example for an Outlook Add-in developed in Java", application = OfficeApplication.Outlook, loadBehavior = LoadBehavior.LoadOnStart)
public class JoaAddin1 extends OutlookAddin {

	
	final IconManager ribbonIcons;

	private final String USER_AGENT = "Mozilla/5.0";
	
	
	public JoaAddin1() {
	    Globals.setThisAddin(this);
	    ribbonIcons = new IconManager(this);
	}
	
	public Dispatch getRibbonImage(IRibbonControl control) {
	    String imageFileName = control.getTag();
	    Dispatch picdisp = ribbonIcons.get(imageFileName);
	    return picdisp;
	}
	
	public void onSurveyButtonClicked(IRibbonControl ribbonControl) {
		
		BackgTask.run(() -> {
			// Obtain Outlook application interface
            com.wilutions.mslib.outlook.Application app = Globals.getThisAddin().getApplication();
			
            Object owner = getApplication().ActiveWindow();
            
            _Inspector inspector = app.ActiveInspector();
            IDispatch item = inspector.getCurrentItem();
            
            if (item.is(MailItem.class)) {
				MailItem mailItem = item.as(MailItem.class);
				String htmlBody = mailItem.getHTMLBody();
				System.out.println(htmlBody);
        	
				String senderName = mailItem.getSenderName();
				if (!"misysmoods@gmail.com".equalsIgnoreCase(senderName)) {
					MessageBox.show(owner, "MiMood Vote Message", "Sorry this email message is not allowed for MiMoods viewing survey. \n Please select corresponding email message sent by MiMoods", (result, ex) -> {
		            });
					return;
				}
				
            }
			
            MessageBox.show(owner, "MiMood Survey", "You choose to view MiMoods Survey", (result, ex) -> {
	            String topicID = "";
	            
	            if (item.is(MailItem.class)) {
					MailItem mailItem = item.as(MailItem.class);
					String htmlBody = mailItem.getHTMLBody();
					System.out.println(htmlBody);
            	
	            	// parse html to retrieve the topic id
					Document doc = Jsoup.parse(htmlBody);
					Element link = doc.select("a").first();
					String linkText = link.text();
					topicID = linkText.substring(1, linkText.length()-1);
					System.out.println(topicID);
	            }
	            
    			try {
    				URL google = new URL("http://man-l5x96ty1:3000/articles/" + topicID);
    				
    				// open a browser
    	            Desktop desktop = Desktop.isDesktopSupported() ? Desktop.getDesktop() : null;
    	            if (desktop != null && desktop.isSupported(Desktop.Action.BROWSE)) {
    	                try {
    	                    desktop.browse(google.toURI());
    	                } catch (Exception e) {
    	                    e.printStackTrace();
    	                }
    	            }
    			} catch (Exception e) {
    				e.printStackTrace();
    			} 
            });
			
			
			
		  });
	}

		public void onVotesButtonClicked(IRibbonControl ribbonControl) {

		    // Event handlers triggered from Outlook should not immediately call
		    // Outlook functions during processing. This can cause a deadlock.
		    // Thus: execute the code in a background thread.
			
		    BackgTask.run(() -> {

		            // Obtain Outlook application interface
		            com.wilutions.mslib.outlook.Application app = Globals.getThisAddin().getApplication();
		            
		            // get current item
		            _Inspector inspector = app.ActiveInspector();
		            IDispatch item = inspector.getCurrentItem();
		            
		            Object owner = getApplication().ActiveWindow();
		            
		            if (item.is(MailItem.class)) {
						MailItem mailItem = item.as(MailItem.class);
						String htmlBody = mailItem.getHTMLBody();
						System.out.println(htmlBody);
						
						String senderName = mailItem.getSenderName();
						
						if (!"misysmoods@gmail.com".equalsIgnoreCase(senderName)) {
							MessageBox.show(owner, "MiMood Vote Message", "Sorry this email message is not allowed for voting. \n Please select corresponding email message sent by MiMoods", (result, ex) -> {
				            });
							return;
						}
						
						// parse html to retrieve the topic id
						Document doc = Jsoup.parse(htmlBody);
						Element link = doc.select("a").first();
						String linkText = link.text();
						String topicID = linkText.substring(1, linkText.length()-1);
						System.out.println(topicID);
						
						// get email account
						String email = "";
			            Accounts accounts = app.getSession().getAccounts();
						_Account account = accounts.Item(1);
						email = account.getSmtpAddress();
						System.out.println(email);
						
			            String vote = ribbonControl.getId();
						System.out.println(vote);
						
						try {
							int errorCode = sendPUT(topicID,email,vote);
							if (errorCode == 409) {
								MessageBox.show(owner, "MiMood Vote Message", "You have already voted for this mood " + ribbonControl.getId(), (result, ex) -> {
					                System.out.println("MessageBox closed by button=" + result);
					            });
								return;
							}
							
						} catch (Exception e) {
							e.printStackTrace();
						}
						
						
			            MessageBox.show(owner, "MiMood Vote Message", "Thank You. You vote " + ribbonControl.getId(), (result, ex) -> {
			                System.out.println("MessageBox closed by button=" + result);
			            });
					}
		            
		            
//		            _Explorer explorer =  app.ActiveExplorer();
//		            Selection selection = explorer.getSelection();
//		            for (int j = 1; j <= selection.getCount(); j++) {
//						IDispatch element = selection.Item(j);
//						if (element.is(MailItem.class)) {
//							MailItem mailItem = element.as(MailItem.class);
//							mailItem.setSubject("HEHEHE");
//						}
//					}

		    });
		}
		
		// HTTP GET request
		private void sendGet() throws Exception {
	 
			String url = "http://www.google.com/search?q=mkyong";
	 
			URL obj = new URL(url);
			HttpURLConnection con = (HttpURLConnection) obj.openConnection();
	 
			// optional default is GET
			con.setRequestMethod("GET");
	 
			//add request header
			con.setRequestProperty("User-Agent", USER_AGENT);
	 
			int responseCode = con.getResponseCode();
			System.out.println("\nSending 'GET' request to URL : " + url);
			System.out.println("Response Code : " + responseCode);
	 
			BufferedReader in = new BufferedReader(
			        new InputStreamReader(con.getInputStream()));
			String inputLine;
			StringBuffer response = new StringBuffer();
	 
			while ((inputLine = in.readLine()) != null) {
				response.append(inputLine);
			}
			in.close();
	 
			//print result
			System.out.println(response.toString());
	 
		}
		
		private int sendPUT(String topicID, String userEmail, String vote) throws Exception {
			Random random = new Random();
	        URL url = new URL("http://man-l5x96ty1:3001/articles/"+ topicID + "/moods");
	        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
	        
	        connection.setRequestMethod("PUT");
	        connection.setDoOutput(true);
	        connection.setRequestProperty("Content-Type", "application/json");
	        connection.setRequestProperty("Accept", "application/json");
	        
	        DataOutputStream wr = new DataOutputStream(connection.getOutputStream());
			wr.writeBytes(String.format("{\"email\":\""+ userEmail +"\",\"mood\":\"" + vote + "\"}", random.nextInt(30), random.nextInt(20)));
			wr.flush();
			wr.close();
			
			int responseCode = connection.getResponseCode();
			
			if (responseCode != 200) {
				return responseCode;
			}
			
			BufferedReader in = new BufferedReader(
			        new InputStreamReader(connection.getInputStream()));
			String inputLine;
			StringBuffer response = new StringBuffer();
	 
			while ((inputLine = in.readLine()) != null) {
				response.append(inputLine);
			}
			in.close();
	 
			//print result
			System.out.println(response.toString());
			return responseCode;
			
		}

	
	
}