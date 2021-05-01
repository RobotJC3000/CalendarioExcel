import java.io.File;
import org.apache.log4j.Logger;
import java.io.IOException;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;


public class LeerExcel {
	
	private static Logger log = Logger.getLogger(LeerExcel.class);
	
	public static void main(String[] args) {

		try {
			
			log.debug("==INICIO==");
			File xlsFile = new File("Calendario.xls");

			Workbook workbook = Workbook.getWorkbook(xlsFile);

			// create writable sheet
			Sheet[] sheets = workbook.getSheets();

			if (sheets != null) {

				for (Sheet sheet : sheets) {
					int rows = sheet.getRows();
					int cols = sheet.getColumns();
					
					for (int row = 0; row < rows; row++) {
						for (int col = 0; col < cols; col++) {
							 //System.out.printf("%15s",sheet.getCell(col,row).getContents());
							 log.info(sheet.getCell(col,row).getContents());
						}
						//System.out.println();
						log.info("");
					}

				}
			}
			workbook.close();
		} catch (BiffException e) {
			// TODO Bloque catch generado automáticamente
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Bloque catch generado automáticamente
			e.printStackTrace();
		}

	}

}
