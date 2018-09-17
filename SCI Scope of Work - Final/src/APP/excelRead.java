package APP;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import com.sun.corba.se.spi.orbutil.threadpool.Work;
import org.apache.poi.ss.usermodel.*;
import org.apache.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelRead {

    private Workbook wb;
    private Workbook wb2;
    private Sheet sh;
    private Sheet sh2;
    private FileInputStream fis;
    private FileOutputStream fos;
    private Row row;
    private Cell cell;



    public ListItems[] items;
    public ArrayList<String> selectedItems = new ArrayList<>();


    private ListItems[] repairs;
    private ListItems[] annualMaint;
    private ListItems[] closingCosts;
    private ListItems Brochures;
    private ListItems[] BuildingCost;
    private ListItems[] Cabinets$Van;
    private ListItems[] buildingSale;
    private ListItems[] campaignBro;
    private ListItems[] ceilings;
    private ListItems[] changeOrder;
    private ListItems[] cleanup;
    private ListItems[] coatings;
    private ListItems[] commissions;
    private ListItems[] concrete;
    private ListItems[] conditionReport;
    private ListItems[] constructDraw;
    private ListItems[] contingency;
    private ListItems[] contractorSpec;
    private ListItems[] craneCharge;
    private ListItems[] doorsntrim;
    private ListItems[] dryin;
    private ListItems[] electrical;
    private ListItems[] engineering;
    private ListItems[] equipRepairs;
    private ListItems[] excavation;
    private ListItems[] exclusions;
    private ListItems[] exteriortrim;
    private ListItems[] finance;
    private ListItems[] floors;
    private ListItems[] freight;
    private ListItems[] furnish;
    private ListItems[] granules;
    private ListItems[] hotmop;
    private ListItems[] hvac;
    private ListItems[] inclusions;
    private ListItems[] insulation;
    private ListItems[] insurance;
    private ListItems[] interiorwalls;
    private ListItems[] labor;
    private ListItems[] landscaping;
    private ListItems[] lead;
    private ListItems[] lw;
    private ListItems[] manuwarranties;
    private ListItems[] masonry;
    private ListItems[] metalRoof;
    private ListItems[] misc;
    private ListItems[] mobilization;
    private ListItems[] notes;
    private ListItems[] paint;
    private ListItems[] paymentdraw;
    private ListItems[] permitRun;
    private ListItems[] genInfo;
    private ListItems[] plansPerms;
    private ListItems[] plumbing;
    private ListItems[] prejob;
    private ListItems[] pressureClean;
    private ListItems[] retention;
    private ListItems[] roofFlash;
    private ListItems[] roofFrame;
    private ListItems[] safety;
    private ListItems[] SCIInspection;
    private ListItems[] SCIWarranties;
    private ListItems[] shingleRoof;
    private ListItems[] siding;
    private ListItems[] sitework;
    private ListItems[] specialty;
    private ListItems[] subContractors;
    private ListItems[] supplied;
    private ListItems[] tearOff;
    private ListItems[] tileRoof;
    private ListItems[] timeMaterial;
    private ListItems[] totalCost;
    private ListItems[] trashRemoval;
    private ListItems[] uniFlex;
    private ListItems[] urethane;
    private ListItems[] wallFrame;
    private ListItems[] windowsTrim;
    private ListItems[] woodWork;
    private ListItems[] wpiTrans;
    public String fileName = null;



    public String name, address, phoneNumber, roofType, _notes;


    public excelRead()
    {}

    class ListItems{
        String active = null;
        String item = null;
        String description = null;

    }


    public void excelWrite() throws Exception
    {
        wb2 = new XSSFWorkbook();
        CreationHelper create = wb2.getCreationHelper();

        String[] columns = {"Name", "Address", "Phone Number", "Roof Type", "Notes", "Line Items"};

        sh2 = wb2.createSheet("Proposal" + name);

        Font headerFont = wb2.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = wb2.createCellStyle();
        headerCellStyle.setFont(headerFont);


        Row headerRow = sh2.createRow(0);

        for(int i = 0; i < columns.length; i++)
        {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        String[] lex = {name, address, phoneNumber, roofType, _notes};

        Row row1 = sh2.createRow(1);
        for(int i = 0; i < lex.length; i++) {
            row1.createCell(i).setCellValue(lex[i]);
        }


        int rowNum = 2;
        for(String s1: selectedItems) {
            Row row2 = sh2.createRow(rowNum++);
            row2.createCell(5).setCellValue(s1);
        }

        for(int i = 0; i < columns.length; i++) {
            sh2.autoSizeColumn(i);
        }

        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\SCIRoof\\Desktop\\Scope of Work " + name + ".xlsx");
        wb2.write(fileOut);
        fileOut.close();

        // Closing the workbook
        wb2.close();

    }


    public void readFile() throws Exception
    {
        items = new ListItems[324];
        for(int ix = 0; ix < items.length; ix++)
        {
            items[ix] = new ListItems();
        }

        fis = new FileInputStream("src/Correct Scope of Work.xlsx");
        wb = WorkbookFactory.create(fis);
        sh = wb.getSheet("Sheet1");
        int numRows = sh.getPhysicalNumberOfRows();
        System.out.println("This is the number of rows in Sheet 1: " + numRows);

        DataFormatter data = new DataFormatter();
        Iterator<Row> rowIterate = sh.rowIterator();
        int count = 1;
        int arrayCount = 1;
        for (Row row: sh) {
            for (Cell cell : row) {
                String cellValue = data.formatCellValue(cell);
                if(count == 1)
                {
                    items[arrayCount].active = cellValue;
                    count++;
                }
                else if(count == 2)
                {
                    items[arrayCount].item = cellValue;
                    count++;
                }
                else if(count == 3)
                {
                    items[arrayCount].description = cellValue;
                    count = 1;
                }
                else if(arrayCount == 15)
                {
                    System.out.println("bop");
                }

            }
            arrayCount++;
        }



        for(int i = 1; i < items.length; i++)
        {
            if(items[i].description.equals("5"))
            {
                items[i].description = "N/A";
            }
            System.out.println("Cell "+ (i) + ") " + items[i].active + "\t" + items[i].item + "\t" + items[i].description);
        }
        organize();
    }

    void organize()
    {
        annualMaint = Arrays.copyOfRange(items, 2,8);
        cleanup =  Arrays.copyOfRange(items, 10,15);
        coatings = Arrays.copyOfRange(items, 16,47);
        dryin = Arrays.copyOfRange(items, 51,55);

        engineering = new ListItems[2];
        engineering[0] = items[55];
        engineering[1] = items[56];

        exclusions = Arrays.copyOfRange(items, 57,61);
        granules = Arrays.copyOfRange(items, 63,67);
        hotmop = Arrays.copyOfRange(items, 67,78);
        inclusions = Arrays.copyOfRange(items, 79,82);
        labor = Arrays.copyOfRange(items, 83,86);
        lw = Arrays.copyOfRange(items, 86,89);
        manuwarranties = Arrays.copyOfRange(items, 89,94);
        metalRoof = Arrays.copyOfRange(items, 95,121);
        misc = Arrays.copyOfRange(items, 122,191);
        plansPerms = Arrays.copyOfRange(items, 194,199);
        prejob = Arrays.copyOfRange(items, 200,204);
        pressureClean = Arrays.copyOfRange(items, 204,211);
        repairs = Arrays.copyOfRange(items, 211,245);
        SCIInspection = Arrays.copyOfRange(items, 247,252);
        SCIWarranties = Arrays.copyOfRange(items, 252,262);
        shingleRoof = Arrays.copyOfRange(items, 262,269);
        sitework = Arrays.copyOfRange(items, 269,271);
        supplied = Arrays.copyOfRange(items, 271,274);
        tearOff = Arrays.copyOfRange(items, 274,288);
        tileRoof = Arrays.copyOfRange(items, 288, 297);
        trashRemoval = Arrays.copyOfRange(items, 299,301);
        uniFlex = Arrays.copyOfRange(items, 301,306);
        urethane = Arrays.copyOfRange(items, 306,312);
        woodWork = Arrays.copyOfRange(items, 312,324);



    }

    public ListItems[] getArray(String x) throws Exception {
        organize();
        if (x.equalsIgnoreCase("Annual Maintenance"))
        {
            return annualMaint;
        }
        else if (x.contains("Clean Up"))
        {
            return cleanup;
        }
        else if (x.equalsIgnoreCase("Coating: All"))
        {
            return coatings;
        }

        else if (x.contains("Dry"))
        {
            return dryin;
        }

        else if (x.contains("Engineering"))
        {
            return engineering;
        }
        else if (x.contains("Exclusions"))
        {
            return exclusions;
        }
        else if (x.contains("Granules"))
        {
            return granules;
        }
        else if (x.contains("Hot"))
        {
            return hotmop;
        }
        else if (x.contains("Inclusions"))
        {
            return inclusions;
        }
        else if (x.contains("Labor"))
        {
            return labor;
        }
        else if (x.contains("LW"))
        {
            return lw;
        }
        else if (x.contains("Warranties"))
        {
            return manuwarranties;
        }
        else if (x.contains("Metal"))
        {
            return metalRoof;
        }
        else if (x.contains("MISC"))
        {
            return misc;
        }
        else if (x.contains("Clean Up"))
        {
            return cleanup;
        }
        else if (x.contains("Plans"))
        {
            return plansPerms;
        }
        else if (x.contains("Pre-Job"))
        {
            return prejob;
        }
        else if (x.contains("Pressure"))
        {
            return pressureClean;
        }
        else if (x.contains("Repair"))
        {
            return repairs;
        }
        else if (x.contains("SCI Inspections"))
        {
            return SCIInspection;
        }
        else if (x.contains("Warranties"))
        {
            return SCIWarranties;
        }
        else if (x.contains("Shingle"))
        {
            return shingleRoof;
        }
        else if (x.contains("Site"))
        {
            return sitework;
        }
        else if (x.contains("Supplied"))
        {
            return supplied;
        }
        else if (x.contains("Tear"))
        {
            return tearOff;
        }
        else if (x.contains("Tile"))
        {
            return tileRoof;
        }
        else if (x.contains("Trash"))
        {
            return trashRemoval;
        }
        else if (x.contains("Uniflex"))
        {
            return uniFlex;
        }
        else if (x.contains("Urethane"))
        {
            return urethane;
        }
        else if (x.contains("Wood"))
        {
            return woodWork;
        }





        return null;
    }
}



