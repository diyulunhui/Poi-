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
InputStream is=new FileInputStream("C:\\Users\\86151\\Desktop\\����2�����ݼ�2-�ո�.xlsx");
//workBook�ǹ�����
XSSFWorkbook workBook=new XSSFWorkbook(is);
//size�Ǹù������ڹ�����ĸ���
int size=workBook.getNumberOfSheets();
//д��ѭ����ÿ����������д���
for(int i=0;i<size;i++)
{
	XSSFSheet sheet=workBook.getSheetAt(i);
	//rowNumber��ʾ�ù������ж�����Ч��
	int rowNumber=sheet.getPhysicalNumberOfRows();
	//д��ѭ����ÿ�н��д���
	for(int rowIndex=0;rowIndex<rowNumber;rowIndex++)
	{//ǰ�漸�е������Ҳ���Ҫ��ѡ������
		if(rowIndex==0||rowIndex==1||rowIndex==2)
		{
			continue;
		}
		//���������ж��ٸ���
		XSSFRow row=sheet.getRow(rowIndex);
		//��ÿһ���н��д���
		for(int cellIndex=0;cellIndex<4;cellIndex++)
		{//cell����һ����Ԫ��
			XSSFCell cell=row.getCell(cellIndex);
			//��������Ԫ������ݣ���Ϊ�����Ԫ���������int�࣬�����������Ǹ�����
			System.out.println(cell.getNumericCellValue());
		}
	}
}
}
}
