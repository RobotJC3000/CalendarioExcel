import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Scanner;
import jxl.write.Number;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * 
 * @author Sr.JC
 * @version 2.0
 */
public class EscribirExcelSegundaVersion {
	/*
	 * Genera una aplicacion que pida un año y genere un calendario (12 meses), el
	 * cual lo guardaremos a un fichero excel, generando una hoja para cada mes.
	 * 
	 * las columnas sabado y domingo seran en negrita y la cabecera de los dias en
	 * otro color
	 * 
	 * Utiliza el Log4j para el control de las trazas
	 */
	static Scanner sn = new Scanner(System.in);
	static SimpleDateFormat sdf = null;
	static Date fechaa = null;
	static Calendar calendar = null;

	public static void main(String[] args) {

		// Inicializamos

		String[] fechasFinales = { "01/01/", "02/01/", "03/01/", "04/01/", "05/01/", "06/01/", "07/01/", "08/01/",
				"09/01/", "10/01/", "11/01/", "12/01/" };

		sdf = new SimpleDateFormat("MM/dd/yyyy");

		// Guardarmos el ArrayList que contiene los dias de la semana & y si es Bisiesto
		ArrayList<Integer[]> diaSemanaBisiestoAux = diaSemanaBisiesto(fechasFinales);

		Integer[] nuumAux = diaSemanaBisiestoAux.get(0); // Array de dias de la semana correspondiente al 1 de cada mes.
															// Devuelve un numero
		Integer[] diasMes = diaSemanaBisiestoAux.get(1); // Array de dias totales de cada mes. Si es bisiesto devolverá
															// Febrero con 29. Enero está en la posición 0.

		for (int i = 0; i < nuumAux.length; i++) { // Traza informativa
			System.out.print(nuumAux[i] + ", ");

		}
		System.out.println();
		for (int i = 0; i < diasMes.length; i++) { // Traza informativa
			System.out.print(diasMes[i] + ", ");
		}

		ArrayList<String[][]> mesesAux = new ArrayList<>();

		// Guardarmos el ArrayList que contiene los tableros de los meses colocados
		mesesAux = imprimeMeses(diasMes, nuumAux);

		// Introducimos en cada hoja de Excel el mes correspondiente
		introduceMesesExcel(mesesAux);

	}

	/**
	 * Nos sirve para resetear data
	 * 
	 * @return data
	 */
	public static String[][] tableroMes() {

		String[][] data = { { "L", "M", "X", "J", "V", "S", "D" }, { "_", "_", "_", "_", "_", "_", "_" },
				{ "_", "_", "_", "_", "_", "_", "_" }, { "_", "_", "_", "_", "_", "_", "_" },
				{ "_", "_", "_", "_", "_", "_", "_" }, { "_", "_", "_", "_", "_", "_", "_" },
				{ "_", "_", "_", "_", "_", "_", "_" } };
		return data;
	}

	/**
	 * Nos sirve para resetear dataBool
	 * 
	 * @return dataBool
	 */
	public static boolean[][] tableroMesBool() {
		boolean[][] dataBool = new boolean[7][7];
		return dataBool;
	}

	/**
	 * Introducimos el ArrayList de los meses en Excel
	 * 
	 * @param mesesAux
	 */
	public static void introduceMesesExcel(ArrayList<String[][]> mesesAux) {

		try {
			String[] fechasFestivas = { "1/1","1/2","1/3","1/4","1/5","1/6","1/7","1/8","1/9", "1/10", "12/25"};
			// Create writable workbook

			WritableWorkbook workbook = Workbook.createWorkbook(new File("Calendario.xls"));

			// Create writable sheet
			WritableSheet[] sheets = new WritableSheet[12]; // Una hoja de excel para cada mes (ENERO, FEBRERO...)
			for (int k = 0; k < sheets.length; k++) {
				String sheetAux = "Hoja";
				String kAux = String.valueOf(k);
				String sheetFinal = sheetAux + kAux;
				sheets[k] = workbook.createSheet(sheetFinal, k);
			}
			WritableFont times10ptBoldUnderline = new WritableFont(WritableFont.TIMES, 10, WritableFont.BOLD, false,
					UnderlineStyle.NO_UNDERLINE);
			WritableFont times10ptBoldUnderline2 = new WritableFont(WritableFont.TIMES, 10, WritableFont.NO_BOLD, false,
					UnderlineStyle.NO_UNDERLINE);

			WritableCellFormat formatoNegrita = new WritableCellFormat(times10ptBoldUnderline);
			formatoNegrita.setBackground(Colour.GOLD);
			formatoNegrita.setAlignment(Alignment.CENTRE);


			WritableCellFormat formatoCabecera = new WritableCellFormat(times10ptBoldUnderline);
			formatoCabecera.setBackground(Colour.AQUA);
			formatoCabecera.setAlignment(Alignment.CENTRE);

			WritableCellFormat formatoCuerpo = new WritableCellFormat(times10ptBoldUnderline2);
			formatoCuerpo.setBackground(Colour.GOLD);
			formatoCuerpo.setAlignment(Alignment.CENTRE);

			WritableCellFormat formatoFestivo = new WritableCellFormat(times10ptBoldUnderline);
			formatoFestivo.setBackground(Colour.GREEN);
			formatoFestivo.setAlignment(Alignment.CENTRE);
			
			WritableCellFormat formatoFestivoLaboral = new WritableCellFormat(times10ptBoldUnderline2);
			formatoFestivoLaboral.setBackground(Colour.GREEN);
			formatoFestivoLaboral.setAlignment(Alignment.CENTRE);

			String[][] month;
			for (int l = 0; l < mesesAux.size(); l++) {

				month = mesesAux.get(l);

				for (int i = 0; i < month.length; i++) {
					for (int j = 0; j < month[0].length; j++) {

						// create a cell at position (i, j) and add to the sheet
						Label label = new Label(j, i, month[i][j]);

						if (i == 0) {
							label.setCellFormat(formatoCabecera);

							sheets[l].addCell(label);

						} else {

							if (!month[i][j].equals("_")) {
								
								int dia = Integer.valueOf(month[i][j]);
								String fechaHoy = l+1 + "/" + dia; // l+1 = MM

								if (esFestivo(fechaHoy, fechasFestivas)) {
									if(j < 5) {
										label.setCellFormat(formatoFestivoLaboral);
									}else {
										label.setCellFormat(formatoFestivo);										
									}
									sheets[l].addCell(label);
									
								} else if (j >= 5) {

									label.setCellFormat(formatoNegrita);
									sheets[l].addCell(label);

								} else {
									label.setCellFormat(formatoCuerpo);
									sheets[l].addCell(label);
								}

							} else {

								label.setCellFormat(formatoCuerpo);
								sheets[l].addCell(label);

							}

						}

					}
				}
			}
			workbook.write();
			workbook.close();
		} catch (WriteException | IOException ex) {
			System.out.println(ex.getMessage());
		}
	}

	public static boolean esFestivo(String fecha, String[] fechasFestivas) {

		for (int i = 0; i < fechasFestivas.length; i++) {

			if (fechasFestivas[i].equals(fecha)) {
				return true;
			}

		}

		return false;

	}

	/**
	 * ==> Conseguimos un array con el dia de la semana correspondiente al 1 de cada
	 * mes && ==> Dependiendo los dias del año introducido nos dirá el array si es
	 * bisiesto o no
	 * 
	 * 
	 * @param fechasFinales
	 */
	public static ArrayList<Integer[]> diaSemanaBisiesto(String[] fechasFinales) {

		System.out.println("Introduce el año a comprobar: ");
		String fechaInicial = "";
		fechaInicial = sn.next();
		ArrayList<Integer[]> diaSemanaBisiesto = new ArrayList<>();
		Integer[] numDoceMeses = new Integer[12];
		Integer[] numDocediasMes = new Integer[12];

		for (int i = 0; i < fechasFinales.length; i++) {
			fechasFinales[i] += fechaInicial;

			calendar = Calendar.getInstance(); // Fecha actual

			try {
				fechaa = sdf.parse(fechasFinales[i]); // Transformamos un String a un Date
			} catch (ParseException e) {
				// TODO Bloque catch generado automáticamente
				e.printStackTrace();
			}

			calendar.setTime(fechaa); // seteamos la fecha actual a la que introduzca el usuario
			numDoceMeses[i] = calendar.get(Calendar.DAY_OF_WEEK);

			if (calendar.getActualMaximum(Calendar.DAY_OF_YEAR) == 365) { // Un año Bisiesto tiene 366 días
				Integer[] diasMes = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
				numDocediasMes = diasMes;

			} else {
				Integer[] diasMes = { 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
				numDocediasMes = diasMes;

			}

		}

		diaSemanaBisiesto.add(numDoceMeses);
		diaSemanaBisiesto.add(numDocediasMes);
		return diaSemanaBisiesto;

	}

	/**
	 * Genera un ArrayList con las hojas de calendario de cada mes dependiendo el
	 * anio introducido
	 * 
	 * @param diasMes Array de dias totales de cada mes.
	 * @param nuumAux Array de dias de la semana correspondiente al 1 de cada mes.
	 *                Devuelve un numero.
	 * @return meses
	 */
	public static ArrayList<String[][]> imprimeMeses(Integer[] diasMes, Integer[] nuumAux) {

		String[][] dataAux;
		boolean[][] dataBoolAux;

		ArrayList<String[][]> meses = new ArrayList<>();
		for (int l = 0; l < diasMes.length; l++) {
			dataAux = tableroMes();
			dataBoolAux = tableroMesBool();
			String sAux = "";
			int numAux = 0, sumDias = 0, numD = nuumAux[l];
			for (int i = 1; i < dataAux.length; i++) {
				for (int j = 0; j < dataAux[0].length; j++) {
					if (numD >= 1 && numD <= 7) {

						if (numD == Calendar.SUNDAY) {
							dataBoolAux[1][6] = true;
						}

						else if (numD == Calendar.MONDAY) {
							dataBoolAux[1][0] = true;
						}

						else if (numD == Calendar.TUESDAY) {
							dataBoolAux[1][1] = true;
						}

						else if (numD == Calendar.WEDNESDAY) {
							dataBoolAux[1][2] = true;
						}

						else if (numD == Calendar.THURSDAY) {
							dataBoolAux[1][3] = true;
						}

						else if (numD == Calendar.FRIDAY) {
							dataBoolAux[1][4] = true;
						} else if (numD == Calendar.SATURDAY) {
							dataBoolAux[1][5] = true;
						}
						numD += 20;
					}

					if (dataBoolAux[i][j]) {
						if (dataAux[i][j] == "_" && dataBoolAux[i][j] && sumDias < diasMes[l]) {

							dataAux[i][j] = "1";

							numAux = Integer.parseInt(dataAux[i][j]);
							sumDias += numAux;
							sAux = String.valueOf(sumDias);
							dataAux[i][j] = sAux;

						}
						if (i < 6 && j == 6)
							dataBoolAux[i + 1][0] = true;
						else if (i <= 6 && j < 6)
							dataBoolAux[i][j + 1] = true;
					}

				}
				System.out.println();
			}
			meses.add(dataAux);
		}
		return meses;

	}

}
