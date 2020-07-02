package ActivityTracker.Model;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class mainClass {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		ActivityTrackerOperation op = new ActivityTrackerOperation();
		op.UpdateInOldExcel();

	}

}
