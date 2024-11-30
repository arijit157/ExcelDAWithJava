import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposePieChart{
    public static void main(String[] args) throws Exception {
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Add sample data for the pie chart
        worksheet.getCells().get("A1").putValue("Category");
        worksheet.getCells().get("A2").putValue("Category A");
        worksheet.getCells().get("A3").putValue("Category B");
        worksheet.getCells().get("A4").putValue("Category C");

        worksheet.getCells().get("B1").putValue("Value");
        worksheet.getCells().get("B2").putValue(30);
        worksheet.getCells().get("B3").putValue(50);
        worksheet.getCells().get("B4").putValue(20);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.getCharts().add(ChartType.PIE, 5, 0, 15, 5);
        Chart chart = worksheet.getCharts().get(chartIndex);

        // Set the chart data range
        chart.getNSeries().add("B2:B4", true);
        chart.getNSeries().setCategoryData("A2:A4");

        // Set chart title
        chart.getTitle().setText("Sample Pie Chart");

        // Save the workbook
        workbook.save("C:\\Users\\10744332\\Downloads\\PieChart.xlsx");

        System.out.println("Pie chart created successfully.");
    }
}