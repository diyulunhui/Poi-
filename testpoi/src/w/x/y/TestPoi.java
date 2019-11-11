package w.x.y;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestPoi {
public static void main(String[] args) throws IOException {
InputStream is=new FileInputStream("C:\\Users\\86151\\Desktop\\附件2：数据集2-终稿.xlsx");
//workBook是工作簿
XSSFWorkbook workBook=new XSSFWorkbook(is);
//size是该工作簿内工作表的个数
int size=workBook.getNumberOfSheets();
//写个循环对每个工作表进行处理
for(int i=0;i<size;i++)
{
	XSSFSheet sheet=workBook.getSheetAt(i);
	//rowNumber表示该工作表有多少有效行
	int rowNumber=sheet.getPhysicalNumberOfRows();
	//写个循环对每行进行处理
	for(int rowIndex=0;rowIndex<rowNumber;rowIndex++)
	{//前面几行的数据我不想要，选择跳过
		if(rowIndex==0||rowIndex==1||rowIndex==2)
		{
			continue;
		}
		//计算这行有多少个列
		XSSFRow row=sheet.getRow(rowIndex);
		//对每一个列进行处理
		for(int cellIndex=0;cellIndex<4;cellIndex++)
		{//cell代表一个单元格
			XSSFCell cell=row.getCell(cellIndex);
			//输出这个单元格的数据，因为这个单元格的数据是int类，所以用下面那个函数
			System.out.println(cell.getNumericCellValue());
		}
	}
}
}
}
