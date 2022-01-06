package Data;

import java.io.File;
import java.io.FileInputStream;
import java.lang.ProcessHandle.Info;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;
import org.json.JSONObject;

import com.mongodb.MongoClient;
import com.mongodb.MongoClientURI;
import com.mongodb.client.MongoCollection;

@Path("/call")
public class Data {
	
	@GET
	@Path("/file")
	@Produces(MediaType.APPLICATION_JSON)

	public static MongoClient getConnection() {
		try {
			return new MongoClient(new MongoClientURI("mongodb://admin:myadminpassword@35.154.199.25:27017"));
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}

	@GET
	@Path("/run")
	@Produces(MediaType.APPLICATION_JSON)
	// @Consumes 
	public static void main (String[] args) {

		try {
			
			JSONObject returnobject = new JSONObject();
			MongoClient client = getConnection();
			MongoCollection<Document> userCollection = client.getDatabase("QuickLookDB").getCollection("nikhil");
			System.out.println("RUNNING");
			

			FileInputStream file = new FileInputStream(new File("C:\\Users\\BALRAM\\Downloads\\Balram JSON File.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Row r1= sheet.getRow(0);
			//HashMap <> map1= new HashMap();
			 r1.cellIterator();
			HashMap <Integer, String> map = new HashMap <Integer, String>();
				Iterator<Cell> top =r1.iterator(); 
				while(top.hasNext()) {
					int i=0;
					map.put(i, r1.toString());
					System.out.print(map.toString() + "\t\t\t");		
					i++;
					return;
				}
			 
			Iterator<Row> rowIterator = sheet.iterator();
			
			
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = (Cell) cellIterator.next();
					// Check the cell type and format accordingly
					switch (cell.getCellType()) {
					
					case STRING:
					System.out.print(map.get(top) + "\t\t\t");			
					Document doc = new Document(map.get(top) ,cell.getStringCellValue());
					userCollection.insertOne(doc);
						
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						break;
					case BOOLEAN:
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						break;
					case _NONE:
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
//					case FORMULA:
//						System.out.print(cell.getNumericCellValue() + "\t\t\t");
//						break;
					
					case BLANK:
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						
						break;
					case ERROR:
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						break;
					default:
					}
				}

				System.out.println("");
			}

			file.close();
			returnobject.put("inserted", "True");
			return;
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}