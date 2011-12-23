import java.io.IOException;

import com.jackbe.utilities.*;


public class TestHTTPClientNTML {

	/**
	 * @param args
	 * @throws Exception 
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException, Exception {
		String //urlStr = "http://jackbe.dyndns.org:7070/sites/gene/_vti_bin/Lists.asmx?WSDL";
			//urlStr = "http://jackbe.dyndns.org:7070/sites/gene/MyFiles/test.xls.zip";
	       urlStr = "http://jackbe.dyndns.org:7070/sites/gene/MyGBP/myDoc.doc.zip";
	       String domain = "sps2010"; // May also be referred as realm
	       String userName = "gene";
	       String password = "pwd@sps2010";
	       //String fileName = "test.xls";
	       String fileName = "";
	       
	       httpClientNTML xx = new httpClientNTML();
	       

	       System.out.println(xx.inputStream2String( xx.getAuthenticatedResponseInputStream(urlStr, domain, userName, password, fileName), "UTF-8"));


	}

}
