import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hwpf.*;
import org.apache.poi.hwpf.extractor.*;


public class ReadWord {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		POIFSFileSystem fs = null;
		try {
			fs = new POIFSFileSystem(new FileInputStream("/Users/americosavinon/Documents/workspace/ApachePOI/src/myDoc.doc"));
			
			HWPFDocument document = new HWPFDocument(fs);
			WordExtractor word = new WordExtractor(document);

			String[] paragraphs = word.getParagraphText();
			
			
			
			for(int i=0; i< paragraphs.length;i++){
				System.out.println( i +") Paragraph: " + paragraphs[i] );
			}

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
