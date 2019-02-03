package org.csvintoexcel;

import org.csvintoexcel.cli.CliArgsParser;
import org.csvintoexcel.csvintoexcel.CsvIntoExcel;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {

		// Parser Command Line Arguments ...
		CliArgsParser cliArgsParser = new CliArgsParser(args);
		
		// Create CsvIntoExcel  instance ...
		CsvIntoExcel csvIntoExcel = new CsvIntoExcel(cliArgsParser);
		
		// Clone excel file Input into Output ... 
		csvIntoExcel.cloneInputIntoOutput();
		
		// Copy Formats/Stlyes to use later ...
		csvIntoExcel.copyFormatStyle();
		
		// Clear contents ...
		csvIntoExcel.cleanContent();
		
		// Copy contents...
		csvIntoExcel.copyContents();
		
		// Save new workbook
		csvIntoExcel.saveWorkbook();


        System.out.println( "Done!" );
    }
}
