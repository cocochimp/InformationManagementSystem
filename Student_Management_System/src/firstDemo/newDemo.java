package firstDemo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.*;
import java.util.ArrayList;
import java.util.Scanner;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class newDemo {
    public static void main(String[] args) throws IOException {
        //创建ArrayList数组用于存储Student对象
        ArrayList<Student> array=new ArrayList<>();

        //读取Excel数据并将合法学生的全部数据存储到ArrayList中
        readExcel(array);

        //打印合法学生的姓名和学号
        printNumber(array);

        //查询功能，通过输入学号、姓名查询学生信息
        studentInformation(array);

        //将数据导进Excel中
        writeExcel(array);

        //将数据导入数据库
        updateData(array);

    }

    //1、读取Excel数据并将合法学生的全部数据存储到ArrayList中
    public static void readExcel(ArrayList<Student> array) throws IOException {
        //创建文件输入流对象
        InputStream ips=new FileInputStream("E:\\项目资料\\project_Java\\Finish_Project\\Student_Management_System\\src\\firstDemo\\lib\\Student.xls");

        //创建Excel对象
        HSSFWorkbook wb=new HSSFWorkbook(ips);

        //创建sheet表对象
        HSSFSheet sheet=wb.getSheetAt(0);

        //获取row行数
        int rownums=sheet.getLastRowNum();

        //创建row行对象
        HSSFRow row=sheet.getRow(1);

        //遍历每一行的数据
        for(int i=2;i<=rownums;i++){
            //获取每一个cell单元格数据
            HSSFCell cell01=row.getCell(0);
            HSSFCell cell02=row.getCell(1);

//            //设置单元格类型
            cell01.setCellType(Cell.CELL_TYPE_STRING);
            cell02.setCellType(Cell.CELL_TYPE_STRING);

            //将cell单元数据类型变为String类型
            String cell1=cell01.getStringCellValue();
            String cell2=cell02.getStringCellValue();

            //创建学生对象，将cell中合法元素加入数组中
            Student s=new Student();
            s.setName(cell1);
            s.setId(cell2);

            //获取学号
            String number=s.getId();

            //设置入学年份
            String startYear=number.substring(0,4);
            s.setStartYear(startYear);

            //设置专业编号及设置班级

            //入学年份后两位
            String startYear34=number.substring(2,4);
            String Nprofession=number.substring(4,9);
            String Nklass=number.substring(9,10);

            switch (Nprofession) {
                case "11621" -> {
                    s.setProfession("计算机科学与技术");
                    s.setKlass("计科1" + startYear34 + Nklass);
                }
                case "11671" -> {
                    s.setProfession("信息系统与信息管理");
                    s.setKlass("信管1" + startYear34 + Nklass);
                }
                case "11672" -> {
                    s.setProfession("物联网工程");
                    s.setKlass("物联1" + startYear34 + Nklass);
                }
                case "11701" -> {
                    s.setProfession("软件工程");
                    s.setKlass("软工1" + startYear34 + Nklass);
                }
                default -> {
                    s.setProfession("信息与计算科学");
                    s.setKlass("信计1" + startYear34 + Nklass);
                }
            }

            //设置班级序号
            String klassNumber=number.substring(10);
            s.setNumber(klassNumber);

            //判断学号是否合法&存储到ArrayList中
            judgeNumber(s,array);

            //获取下一行的数据
            row=sheet.getRow(i);
        }

        //关闭字节输入流
        ips.close();
    }

    //2、打印合法学生的姓名和学号
    public static void printNumber(ArrayList<Student> array){
        //控制台输出合法成员
        System.out.println("Excel表中合法的成员有：");

        //学生人数
        int sum=0;

        //增强for循环遍历合法成员
        for(Student sb:array){
            System.out.println(sb.getName()+","+ sb.getId());
            sum++;
        }
        System.out.println("合法学生一共有"+sum+"名");
    }

    //3、查询功能，通过输入学号、姓名查询学生信息
    public static void studentInformation(ArrayList<Student> array) {

        while(true){
            //创建键盘输入对象
            Scanner sc = new Scanner(System.in);
            //输入你要查找的名字
            System.out.println("\r\n"+"一、你要查找的学生姓名或学号：[输入0退出查询功能]");
            String nameNumber = sc.nextLine();

            if(nameNumber.equals("0")){
                System.out.println("谢谢使用！欢迎下次再用！");
                break;
            }

            //通过姓名或者学号查找学生信息
            for (int i = 0; i < array.size(); i++) {
                Student st = array.get(i);
                if (nameNumber.equals(st.getId()) || nameNumber.equals(st.getName())) {
                    System.out.println("匹配成功！" + "\n"+"该学生的信息是：");
                    System.out.println(st.getName()+","+st.getId()+","+st.getStartYear()+","+st.getProfession()+","+st.getKlass()+","+st.getNumber());
                    break;
                } else if (i == array.size() - 1) {
                    Student sd = array.get(i);
                    if (!nameNumber.equals(sd.getId()) || !nameNumber.equals(sd.getName())) {
                        System.out.println("匹配失败！"+"\n"+"请重新匹配！");
                        break;
                    }
                }

            }
        }

    }

    //4、输出17级学生的Excel数据
    public static void writeExcel17(ArrayList<Student> array) throws IOException {
        //创建HSSFWorkbook对象
        HSSFWorkbook wb = new HSSFWorkbook();
        //创建Sheet页表
        HSSFSheet sheet = wb.createSheet("sheet0");
        //创建row行数
        HSSFRow row1 = sheet.createRow(0);
        //创建单元格
        HSSFCell cell11=row1.createCell(0);
        HSSFCell cell12=row1.createCell(1);
        //设置单元格的信息
        cell11.setCellValue("姓名");
        cell12.setCellValue("学号");

        //设置初始行数
        int rowNumber17=1;

        //将学生数据进行判断并写入HSSFWorkbook对象中
        for (Student s : array) {
            //获取学号中的3-4为作为年级的评定标准
            String year = s.getStartYear();
            String year34 = year.substring(2);

            if (year34.equals("17")) {
                //创建row行数
                HSSFRow row17 = sheet.createRow(rowNumber17);
                //创建单元格
                HSSFCell cell01 = row17.createCell(0);
                HSSFCell cell02 = row17.createCell(1);
                //设置单元格的信息
                cell01.setCellValue(s.getName());
                cell02.setCellValue(s.getId());
                rowNumber17++;
            }
        }

        //创建字节输出流对象
        FileOutputStream output17=new FileOutputStream("E:\\项目资料\\project_Java\\Finish_Project\\Student_Management_System\\src\\firstDemo\\lib\\2017.xlsx");

        //将HSSFWorkbook对象内容写入Excel中
        wb.write(output17);

        //释放资源
        output17.flush();

        System.out.println("成功将17级合法学生的数据输出到Excel中！");
    }

    //5、输出18级学生的Excel数据
    public static void writeExcel18(ArrayList<Student> array) throws IOException {
        //创建HSSFWorkbook对象
        HSSFWorkbook wb = new HSSFWorkbook();
        //创建Sheet页表
        HSSFSheet sheet = wb.createSheet("sheet0");
        //创建row行数
        HSSFRow row1 = sheet.createRow(0);
        //创建单元格
        HSSFCell cell11=row1.createCell(0);
        HSSFCell cell12=row1.createCell(1);
        //设置单元格的信息
        cell11.setCellValue("姓名");
        cell12.setCellValue("学号");

        //设置初始行数
        int rowNumber18=1;

        //将学生数据进行判断并写入HSSFWorkbook对象中
        for (Student s : array) {
            //获取学号中的3-4为作为年级的评定标准
            String year = s.getStartYear();
            String year34 = year.substring(2);

            if (year34.equals("18")) {
                //创建row行数
                HSSFRow row18 = sheet.createRow(rowNumber18);
                //创建单元格
                HSSFCell cell01 = row18.createCell(0);
                HSSFCell cell02 = row18.createCell(1);
                //设置单元格的信息
                cell01.setCellValue(s.getName());
                cell02.setCellValue(s.getId());
                rowNumber18++;
            }
        }

        //创建字节输出流对象
        FileOutputStream output18=new FileOutputStream("E:\\项目资料\\project_Java\\Finish_Project\\Student_Management_System\\src\\firstDemo\\lib\\2018.xlsx");

        //将HSSFWorkbook对象内容写入Excel中
        wb.write(output18);

        //释放资源
        output18.flush();

        System.out.println("成功将18级合法学生的数据输出到Excel中！");
    }

    //6、输出19级学生的Excel数据
    public static void writeExcel19(ArrayList<Student> array) throws IOException {
        //创建HSSFWorkbook对象
        HSSFWorkbook wb = new HSSFWorkbook();
        //创建Sheet页表
        HSSFSheet sheet = wb.createSheet("sheet0");
        //创建row行数
        HSSFRow row1 = sheet.createRow(0);
        //创建单元格
        HSSFCell cell11=row1.createCell(0);
        HSSFCell cell12=row1.createCell(1);
        //设置单元格的信息
        cell11.setCellValue("姓名");
        cell12.setCellValue("学号");

        //设置初始行数
        int rowNumber19=1;

        //将学生数据进行判断并写入HSSFWorkbook对象中
        for (Student s : array) {
            //获取学号中的3-4为作为年级的评定标准
            String year = s.getStartYear();
            String year34 = year.substring(2);

            if (year34.equals("19")) {
                //创建row行数
                HSSFRow row19 = sheet.createRow(rowNumber19);
                //创建单元格
                HSSFCell cell01 = row19.createCell(0);
                HSSFCell cell02 = row19.createCell(1);
                //设置单元格的信息
                cell01.setCellValue(s.getName());
                cell02.setCellValue(s.getId());
                rowNumber19++;
            }
        }

        //创建字节输出流对象
        FileOutputStream output19=new FileOutputStream("E:\\项目资料\\project_Java\\Finish_Project\\Student_Management_System\\src\\firstDemo\\lib\\2019.xlsx");

        //将HSSFWorkbook对象内容写入Excel中
        wb.write(output19);

        //释放资源
        output19.flush();

        System.out.println("成功将19级合法学生的数据输出到Excel中！");
    }

    //7、将各年级学生的数据放入Excel中
    public static void writeExcel(ArrayList<Student> array) throws IOException {
        System.out.println("\r\n"+"二、你是否要将合法学生数据放进Excel中保存？[是：1]");
        //输出各个年级的Excel数据
        Scanner sc=new Scanner(System.in);
        int s = sc.nextInt();

        if(s==1){
            writeExcel17(array);
            writeExcel18(array);
            writeExcel19(array);
        }

        System.out.println("谢谢使用！欢迎下次再用！");
    }

    //8、判断学号是否合法
    public static void judgeNumber(Student s,ArrayList<Student> array){
        //设置学号合法规则
        String numberRude01 = "20(17|18|19|20)11(621|671|672|701|921)[1-9]\\d{2}";
        String numberRude02 = "\\d{12}";

        boolean flag01 =s.getId().matches(numberRude01);
        boolean flag02 =s.getId().matches(numberRude02);

        if(flag01 && flag02){
            array.add(s);
        }
    }

    //9、删除一条原有数据
    public static void delete(String id){

        Connection con=null;
        PreparedStatement st=null;

        try {
            //3、获取数据库对象
            con = JdbcUtils.getConnection();

            //区别
            //4、执行SQL语句
            //使用占位符代替参数
            String sql="delete from all_student where id=?";

            //预编译，先写sql，不执行
            st=con.prepareStatement(sql);
            st.setString(1,id);

            //执行
            st.executeUpdate();


        } catch (SQLException e) {
            e.printStackTrace();
        }finally{
            //6、释放资源
            JdbcUtils.release(con,st,null);
        }
    }

    //10、插入一条新数据
    public static void insert(String id, String name,String startYear, String profession, String klass, String number){

        Connection con=null;
        PreparedStatement st=null;

        try {
            //3、获取数据库对象
            con = JdbcUtils.getConnection();

            //区别
            //4、执行SQL语句
            //使用占位符代替参数
            String sql="insert into all_student(id,name,startYear,profession,klass,number) values(?,?,?,?,?,?)";

            //预编译，先写sql，不执行
            st=con.prepareStatement(sql);
            st.setString(1,id);
            st.setString(2,name);
            st.setString(3,startYear);
            st.setString(4,profession);
            st.setString(5,klass);
            st.setString(6,number);

            //执行
            st.executeUpdate();

        } catch (SQLException e) {
            e.printStackTrace();
        }finally{
            //6、释放资源
            JdbcUtils.release(con,st,null);
        }
    }

    //11、显示所有数据
    public static void login(){

        Connection con=null;
        PreparedStatement st=null;
        ResultSet rs=null;

        try {
            //3、获取数据库对象
            con = JdbcUtils.getConnection();

            //区别
            //4、执行SQL语句
            //使用占位符代替参数
            String sql="select * from all_student";

            //预编译，先写sql，不执行
            st=con.prepareStatement(sql);

            //执行
            rs = st.executeQuery();

            System.out.println("all_student表中数据:");

            while(rs.next()){
                System.out.print(rs.getString(1)+"\t");
                System.out.print(rs.getString(2)+"\t");
                System.out.print(rs.getString(3)+"\t");
                System.out.print(rs.getString(4)+"\t  ");
                System.out.print(rs.getString(5)+"\t  ");
                System.out.println(rs.getString(6));
            }

        } catch (SQLException e) {
            e.printStackTrace();
        }finally{
            //6、释放资源
            JdbcUtils.release(con,st,rs);
        }
    }

    //12、控制数据库操作
    public static void updateData(ArrayList<Student> array){
        System.out.println("\r\n"+"三、输入你要实现的功能：");
        System.out.println("1、更新数据库数据");
        System.out.println("2、显示数据库现有数据");
        System.out.println("【输入其他退出功能！！！】");
        Scanner sc=new Scanner(System.in);

        while(true){
            System.out.println("\r\n"+"输入你要实现的功能值:");
            int s=sc.nextInt();
            if(s==1){
                for (Student st : array) {
                    //删除原有数据
                    delete(st.getId());

                    //将数据导入数据库
                    insert(st.getId(), st.getName(), st.getStartYear(), st.getProfession(), st.getKlass(), st.getNumber());
                }
                System.out.println("数据库数据更新完毕！！！");
            }else if(s==2){
                //显示数据库所有数据
                login();

            }else{
                System.out.println("谢谢使用！欢迎下次再用！");
                break;
            }
        }

    }
}


