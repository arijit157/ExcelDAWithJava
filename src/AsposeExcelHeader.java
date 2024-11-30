import com.aspose.cells.*;

public class AsposeExcelHeader {
    public static void main(String[] args) throws Exception {
        Workbook workbook = new Workbook();

        Worksheet workSheet = workbook.getWorksheets().get(0);

        Cell cellH1 =workSheet.getCells().get("A1");
        Cell cellH2 =workSheet.getCells().get("B1");

        cellH1.putValue("Product Name");
        cellH2.putValue("No.of APIs");

        Style style1 = cellH1.getStyle();
        Style style2 = cellH2.getStyle();

        Font font1 = style1.getFont();
        Font font2 = style2.getFont();

        font1.setBold(true);
        font2.setBold(true);

        style1.setHorizontalAlignment(TextAlignmentType.CENTER);
        style2.setHorizontalAlignment(TextAlignmentType.CENTER);

        cellH1.setStyle(style1);
        cellH2.setStyle(style2);

        workbook.save("C:\\Users\\10744332\\Downloads\\app_heading.xlsx");
    }
}
