import com.aspose.cells.*;

import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

public class Main {
    public static void main(String[] args) throws Exception {
        Workbook workbook = new Workbook("C:\\Users\\10744332\\Downloads\\app.xlsx");

        Worksheet workSheet = workbook.getWorksheets().get(0);

        Range range1 = workSheet.getCells().createRange("A2:A42");

        Range range2 = workSheet.getCells().createRange("D2:D42");

        Set<String> uniqueProdNames = new HashSet<>();

        Set<String> uniqueStatusCode = new HashSet<>();

        //getting the unique product names
        for(int i=0; i<range1.getRowCount(); i++){
            Cell cell = range1.get(i, 0);
            uniqueProdNames.add(cell.getStringValue());
        }

        //getting the unique HTTP status codes
        for(int j=0; j<range2.getRowCount(); j++){
            Cell cell = range2.get(j, 0);
            uniqueStatusCode.add(cell.getStringValue());
        }

        System.out.println(uniqueStatusCode);

        List<String> productNames = uniqueProdNames.stream().toList();

        List<String> statusCodes = uniqueStatusCode.stream().toList();

        Workbook newWorkBook = new Workbook();

        Worksheet newWorkSheet = newWorkBook.getWorksheets().get(0);

        Cell cellH1 =newWorkSheet.getCells().get("A1");

        Cell cellH2 = newWorkSheet.getCells().get("B1");

        Cell cellH3 = newWorkSheet.getCells().get("C1");

        Cell cellH4 = newWorkSheet.getCells().get("D1");

        Cell cellH5 = newWorkSheet.getCells().get("E1");

        cellH1.putValue("Product Name");

        cellH2.putValue("No. of API");

        cellH3.putValue("500 Status Code Count");

        cellH4.putValue("401 Status Code Count");

        cellH5.putValue("404 Status Code Count");

        Style style = cellH1.getStyle();

        Font font = style.getFont();

        font.setBold(true);

        style.setHorizontalAlignment(TextAlignmentType.CENTER);

        cellH1.setStyle(style);
        cellH2.setStyle(style);
        cellH3.setStyle(style);
        cellH4.setStyle(style);
        cellH5.setStyle(style);

        Iterator<String> iterator = uniqueProdNames.iterator();

        int i=0;

        while(iterator.hasNext()){
            String prodName = iterator.next();
            newWorkSheet.getCells().get(i+1, 0).putValue(prodName);
            i++;
        }

        ////////////////////// To Count the number of API Operations present in each products //////////////////////////////////
        for(int j=0; j<productNames.size(); j++){
            Cell cell = workSheet.getCells().get("U2");

            cell.setFormula("=COUNTIF(A:A, \""+productNames.get(j)+"\")");

            CalculationOptions options = new CalculationOptions();

            options.setIgnoreError(true);

            workbook.calculateFormula(options);

            int totalNoOfAPIS = cell.getIntValue();

            System.out.println("Total APIs: "+totalNoOfAPIS);

            Cell cell2 = newWorkSheet.getCells().get("B"+(j+2));

            cell2.putValue(totalNoOfAPIS);
        }
        ////////////////////////////////////////////////////////////////////////////////////////////////

        String nextColName = "B";

        for(int k=0; k<statusCodes.size(); k++){

            int colIndex = CellsHelper.columnNameToIndex(nextColName);
            nextColName = CellsHelper.columnIndexToName(colIndex+1);

            for(int j=0; j<productNames.size(); j++) {

                Cell cell = workSheet.getCells().get("U2");

                cell.setFormula("=COUNTIFS(A:A, \"" + productNames.get(j) + "\", D:D, \""+statusCodes.get(k)+"\")");

                CalculationOptions options = new CalculationOptions();

                options.setIgnoreError(true);

                workbook.calculateFormula(options);

                int totalNumberOfStatusCode = cell.getIntValue();

                Cell cell2 = newWorkSheet.getCells().get(nextColName + (j + 2));

                cell2.putValue(totalNumberOfStatusCode);
            }
        }

        //////////////////////// Count of the 500 status code in each API //////////////////////////////

//        for(int j=0; j<productNames.size(); j++) {
//            Cell cell = workSheet.getCells().get("U2");
//
//            cell.setFormula("=COUNTIFS(A:A, \"" + productNames.get(j) + "\", D:D, \"500\")");
//
//            CalculationOptions options = new CalculationOptions();
//
//            options.setIgnoreError(true);
//
//            workbook.calculateFormula(options);
//
//            int totalNumberOfStatusCode = cell.getIntValue();
//
//            Cell cell2 = newWorkSheet.getCells().get("C" + (j + 2));
//
//            cell2.putValue(totalNumberOfStatusCode);
//        }
        ////////////////////////////////////////////////////////////////////////////////////////////////

        //////////////////////// Count of the 401 status code in each API //////////////////////////////
//        for(int j=0; j<productNames.size(); j++) {
//            Cell cell = workSheet.getCells().get("U2");
//
//            cell.setFormula("=COUNTIFS(A:A, \"" + productNames.get(j) + "\", D:D, \"401\")");
//
//            CalculationOptions options = new CalculationOptions();
//
//            options.setIgnoreError(true);
//
//            workbook.calculateFormula(options);
//
//            int totalNumberOfStatusCode = cell.getIntValue();
//
//            Cell cell2 = newWorkSheet.getCells().get("D" + (j + 2));
//
//            cell2.putValue(totalNumberOfStatusCode);
//        }
        ///////////////////////////////////////////////////////////////////////////////////////////////

        //////////////////////// Count of the 404 status code in each API //////////////////////////////
//        for(int j=0; j<productNames.size(); j++) {
//            Cell cell = workSheet.getCells().get("U2");
//
//            cell.setFormula("=COUNTIFS(A:A, \"" + productNames.get(j) + "\", D:D, \"404\")");
//
//            CalculationOptions options = new CalculationOptions();
//
//            options.setIgnoreError(true);
//
//            workbook.calculateFormula(options);
//
//            int totalNumberOfStatusCode = cell.getIntValue();
//
//            Cell cell2 = newWorkSheet.getCells().get("E" + (j + 2));
//
//            cell2.putValue(totalNumberOfStatusCode);
//        }
        ////////////////////////////////////////////////////////////////////////////////////

//        createBarChart1(newWorkBook);
//        createBarChart2(newWorkBook);
//        createBarChart3(newWorkBook);

        createBarChart(newWorkBook, "500", "C1:C7", "A1:A7");
        createBarChart(newWorkBook, "404", "E1:E7", "A1:A7");
        createBarChart(newWorkBook, "401", "D1:D7", "A1:A7");

        createPieChart(newWorkBook, "500", "C1:C7", "A1:A7");
        createPieChart(newWorkBook, "404", "E1:E7", "A1:A7");
        createPieChart(newWorkBook, "401", "D1:D7", "A1:A7");

        newWorkBook.save("C:\\Users\\10744332\\Downloads\\app_updated.xlsx");

        System.out.println("Data updated successfully!");
    }

    //function for creating the chart
//    public static void createBarChart1(Workbook workbook) throws Exception {
//        Worksheet worksheet = workbook.getWorksheets().get(0);
//        int chartIndex = worksheet.getCharts().add(ChartType.BAR, 5, 0, 20, 10);
//        Chart chart = worksheet.getCharts().get(chartIndex);
//        chart.getNSeries().add("C1:C7", true);
//        chart.getNSeries().setCategoryData("A1:A7");
//        chart.getTitle().setText("Product Name VS 500 Status Code");  //setting the chart title
//        workbook.save("C:\\Users\\10744332\\Downloads\\ProdName_500_BarChart.xlsx");
//        System.out.println("Bar chart created successfully.");
//    }

//    public static void createBarChart2(Workbook workbook) throws Exception {
//        Worksheet worksheet = workbook.getWorksheets().get(0);
//        int chartIndex = worksheet.getCharts().add(ChartType.BAR, 5, 0, 20, 10);
//        Chart chart = worksheet.getCharts().get(chartIndex);
//        chart.getNSeries().add("E1:E7", true);
//        chart.getNSeries().setCategoryData("A1:A7");
//        chart.getTitle().setText("Product Name VS 404 Status Code");  //setting the chart title
//        workbook.save("C:\\Users\\10744332\\Downloads\\ProdName_404_BarChart.xlsx");
//        System.out.println("Bar chart created successfully.");
//    }

//    public static void createBarChart3(Workbook workbook) throws Exception {
//        Worksheet worksheet = workbook.getWorksheets().get(0);
//        int chartIndex = worksheet.getCharts().add(ChartType.BAR, 5, 0, 20, 10);
//        Chart chart = worksheet.getCharts().get(chartIndex);
//        chart.getNSeries().add("D1:D7", true);
//        chart.getNSeries().setCategoryData("A1:A7");
//        chart.getTitle().setText("Product Name VS 401 Status Code");  //setting the chart title
//        workbook.save("C:\\Users\\10744332\\Downloads\\ProdName_401_BarChart.xlsx");
//        System.out.println("Bar chart created successfully.");
//    }

        public static void createPieChart(Workbook workBook, String statusCodeName, String xAxisRange, String yAxisRange) throws Exception {
            Worksheet worksheet = workBook.getWorksheets().get(0);
            int chartIndex = worksheet.getCharts().add(ChartType.PIE, 5, 0, 20, 10);
            Chart chart = worksheet.getCharts().get(chartIndex);
            chart.getNSeries().add(xAxisRange, true);  //x-axis
            chart.getNSeries().setCategoryData(yAxisRange);  //y-axis
            chart.getTitle().setText("Product Name VS " + statusCodeName + " Status Code");  //setting the chart title
//            workBook.save("C:\\Users\\10744332\\Downloads\\ProdName_" + statusCodeName + "_PieChart.xlsx");
            System.out.println("Pie chart created successfully.");
        }

    public static void createBarChart(Workbook workBook, String statusCodeName, String xAxisRange, String yAxisRange) throws Exception {
        Worksheet worksheet = workBook.getWorksheets().get(0);
        int chartIndex = worksheet.getCharts().add(ChartType.BAR, 5, 0, 20, 10);
        Chart chart = worksheet.getCharts().get(chartIndex);
        chart.getNSeries().add(xAxisRange, true);  //x-axis
        chart.getNSeries().setCategoryData(yAxisRange);  //y-axis
        chart.getTitle().setText("Product Name VS "+statusCodeName+" Status Code");  //setting the chart title
//        workBook.save("C:\\Users\\10744332\\Downloads\\ProdName_"+statusCodeName+"_BarChart.xlsx");
        System.out.println("Bar chart created successfully.");
    }
}