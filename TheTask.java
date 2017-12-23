import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;


public class TheTask {

    public static void main(String[] Arg) throws IOException {

        Collection<Schedule> Collection;
        int CountMonth;
        double Amount;
        double Percent;
        HSSFWorkbook workBook = new HSSFWorkbook(); // Создание книги
        Sheet sheet = workBook.createSheet("Лист 1"); // Создание листа

        Scanner input = new Scanner(System.in);
        System.out.println("Формула расчета аннуитетного кредита");
        System.out.println("====================================");
        System.out.println("<Кол-во месяцев> <Сумма> <%-ая ставка за год>");

        CountMonth = input.nextInt();
        Amount = input.nextDouble();
        Percent = input.nextDouble();
        Collection=InitColl(CountMonth,Amount,Percent);

        for (Schedule test : Collection)
       {
           System.out.println(test.getDate() + ", сумма платежа: " + test.getOd() + ", остаток по долгу: " + test.getPvc());
       }

        int rowNum = 0;
        Cell cell;
        Row row;
        HSSFCellStyle style = createStyleForTitle(workBook);

        row = sheet.createRow(rowNum);
        // Date
        cell = row.createCell(0, CellType.STRING);
        cell.setCellValue("Date");
        cell.setCellStyle(style);
        // Od
        cell = row.createCell(1, CellType.STRING);
        cell.setCellValue("Od");
        cell.setCellStyle(style);
        // Pvc
        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("Pvc");
        cell.setCellStyle(style);

        for (Schedule emp : Collection) {
            rowNum++;
            row = sheet.createRow(rowNum);

            // Date (A)
            cell = row.createCell(0, CellType.STRING);
            cell.setCellValue(emp.getDate());
            // Od (B)
            cell = row.createCell(1, CellType.STRING);
            cell.setCellValue(emp.getOd().toString());
            // Pvc (C)
            cell = row.createCell(2, CellType.STRING);
            cell.setCellValue(emp.getPvc().toString());
        }

        FileOutputStream fos = new FileOutputStream("C:\\Users\\Normandy\\IdeaProjects\\JavaTask1\\book.xlsx"); // Создать поток записи
        workBook.write(fos); // Создание книги
        fos.close(); // Закрытие потока записи





}
    private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        return style;
    }


    public static Collection<Schedule> InitColl(int _CountMonth, double _Amount, double _Percent) {

        Calendar CalendarForProcent = (Calendar) Calendar.getInstance().clone();//Календарь для процентов
        int M = CalendarForProcent.get(Calendar.MONTH) + 1;//+1 тк нумерация месяцев начинается с 0
        int Y = CalendarForProcent.get(Calendar.YEAR);
        Collection<Schedule> Collection = new ArrayList<>(); // Коллекция для подсчета каждого месяца
        BigDecimal SumWithProcent= new BigDecimal(_Amount);// сумма кредита

        _Percent = _Percent / 12 / 100;//Годовая ставка, делим на 12

        BigDecimal Od = new BigDecimal(_Amount*(_Percent*Math.pow(1+_Percent,_CountMonth)/(Math.pow(1+_Percent,_CountMonth)-1))).setScale(4, RoundingMode.HALF_UP);//сумма ежемесячного платежа
        Od=Od.setScale(2,BigDecimal.ROUND_HALF_UP);

        CalendarForProcent.set(Y, M, 1);
        int day_M = CalendarForProcent.getActualMaximum(Calendar.DAY_OF_MONTH);//Количество дней в месяце
        int day_Y = CalendarForProcent.getActualMaximum(Calendar.DAY_OF_YEAR);//Количество дней в году

        BigDecimal PayProcent = new BigDecimal((_Percent*12/day_Y*day_M)*_Amount);//Начисленные проценты в этом месяце
        PayProcent=PayProcent.setScale(2,BigDecimal.ROUND_HALF_UP);

        Collection.add(new Schedule(DateToString(M, Y),Od,SumWithProcent=SumWithProcent.subtract(Od.subtract(PayProcent))));

        for (int i=2; i<=_CountMonth;i++)
        {
            M++;
            if (M>12) {M=1; Y++;}

            CalendarForProcent.set(Y, M, 1);
            day_M = CalendarForProcent.getActualMaximum(Calendar.DAY_OF_MONTH);
            day_Y = CalendarForProcent.getActualMaximum(Calendar.DAY_OF_YEAR);

            BigDecimal SUM = new BigDecimal(_Percent*12/day_Y*day_M);
            PayProcent = SumWithProcent.multiply((SUM));
            PayProcent=PayProcent.setScale(2,BigDecimal.ROUND_HALF_UP);

            Collection.add(new Schedule(DateToString(M, Y),Od,SumWithProcent=SumWithProcent.subtract(Od.subtract(PayProcent))));
        }

        return Collection;
    }

    public static String DateToString(int month, int year)
    {
        String date = new String();

        switch(month)
        {
            case 1:
                date += "Январь";
                break;
            case 2:
                date += "Февраль";
                break;
            case 3:
                date += "Март";
                break;
            case 4:
                date += "Апрель";
                break;
            case 5:
                date += "Май";
                break;
            case 6:
                date += "Июнь";
                break;
            case 7:
                date += "Июлб";
                break;
            case 8:
                date += "Август";
                break;
            case 9:
                date += "Сентябрь";
                break;
            case 10:
                date += "Октябрь";
                break;
            case 11:
                date += "Ноябрь";
                break;
            case 12:
                date += "Декабрь";
                break;
            default:
                System.out.println("Ошибка считывания месяца");
                break;
        }

        date += " "+year;
        return date;
        }
    }


class Schedule {

    private String Date; // Месяц + Год выплаты
    private BigDecimal Od;   // Выплата за месяц
    private BigDecimal Pvc;  // Конечная сумма на выплату

    public Schedule(String _Date, BigDecimal _Od, BigDecimal _Pvc) {
        Date = _Date;
        Od = _Od;
        Pvc = _Pvc;
    }

    public Schedule() {
        Date = null;
        Od = Pvc = null;
    }

    public String getDate() {
        return Date;
    }

    public void setDate(String _Date) {
        this.Date = _Date;
    }

    public BigDecimal getOd() {
        return Od;
    }

    public void setOd(BigDecimal _Od) {
        this.Od = _Od;
    }

    public BigDecimal getPvc() {
        return Pvc;
    }

    public void setPvc(BigDecimal _Pvc) {
        Pvc = _Pvc;
    }

}


