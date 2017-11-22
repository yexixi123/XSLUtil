import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;

/**
 * Created by Think on 2017/7/27.
 */
public class XSL {

    public static void main(String[] args) throws Exception {
        try {
            XSL xsl = new XSL();
            ArrayList<String> txtDate = getTXTDate("D:\\XSL\\src\\main\\java\\xsl.txt");
            String[] colName = txtDate.get(1).split(",");//列名集合
            String[] connDate = txtDate.get(2).split("###");//连接数据库的 url，username,password

            ArrayList<HashMap<String, String>> xsl1 = xsl.getXSL1("D:\\下载\\数据\\税种.xls", colName);

            Connection connection = getConnection(connDate);
            xsl.queryXSL(txtDate.get(0), colName, xsl1, connection);

        } catch (Exception e1) {
            e1.printStackTrace();
        }

    }

    private ArrayList<HashMap<String, String>> getXSL1(String file, String[] colName) throws IOException {
        if (file == null || file == "") {
            System.err.println("数据文件不存在");
        }
        File file1 = new File(file);
        InputStream is = new FileInputStream(file1);
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
        //获取第一个工作簿
        HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
        if (hssfSheet == null) {
            return new ArrayList<HashMap<String, String>>();
        }
        //获得总列数
        int coloumNum = hssfSheet.getRow(0).getPhysicalNumberOfCells();
        //获得总行数
        int rowNum = hssfSheet.getLastRowNum();

        ArrayList<HashMap<String, String>> arrayList = new ArrayList<HashMap<String, String>>();
        HashMap<String, String> stringStringHashMap = null;

        // 获取当前工作薄的每一行
        for (int i = 0; i <= rowNum; i++) {
            HSSFRow hssfRow = hssfSheet.getRow(i);
            if (hssfRow != null) {
                stringStringHashMap = new HashMap<String, String>();
                for (int j = 0; j < coloumNum; j++) {
                    HSSFCell one = hssfRow.getCell(j);
                    //读取第一列数据
                    if (one != null) {
                        String stringCellValue = one.getStringCellValue();
                        stringStringHashMap.put(colName[j], stringCellValue);
                    }

                }
            }
            arrayList.add(stringStringHashMap);
        }

        return arrayList;
    }

    /**
     * 获取xls 数据插入数据库
     *
     * @param sql        sql 语句
     * @param coloumName 列名
     * @param xsl        数据
     * @param connection 数据库连接
     */
    public void queryXSL(String sql, String[] coloumName, ArrayList<HashMap<String, String>> xsl, Connection connection) {
        PreparedStatement psts = null;
        int count = 0;
        try {
            connection.setAutoCommit(false); // 设置手动提交
            psts = connection.prepareStatement(sql);
            if (xsl.size() > 0) {

                for (HashMap map : xsl) {
                    int i = 1;
                    for (String name : coloumName) {
                        String coloum = (String) map.get(name);
                        psts.setObject(i++, coloum);
                    }
                    psts.addBatch();          // 加入批量处理
                    count++;
                }
                psts.executeBatch(); // 执行批量处理
                connection.commit();  // 提交
                System.out.println("All down : " + count);
                psts.close();
                connection.close();
            }
        } catch (Exception e1) {
            e1.printStackTrace();
        }

    }

    /**
     * 获取数据库连接
     *
     * @param connDate
     * @return
     * @throws Exception
     */
    public static Connection getConnection(String[] connDate) throws Exception {
        Connection con = null;
        try {
            Class.forName(connDate[0]);
            con = DriverManager.getConnection(connDate[1], connDate[2], connDate[3]);
        } catch (SQLException se) {

            System.out.println("数据库连接失败！");
            se.printStackTrace();
        }
        return con;
    }

    /**
     * 获取 txt文件的配置数据
     *
     * @return
     */
    public static ArrayList<String> getTXTDate(String filePath) {
        if (filePath == null || filePath == "") {
            System.err.println("配置文件不存在");
        }
        BufferedReader reader = null;
        ArrayList<String> strings = new ArrayList<String>();
        try {
            File file = new File(filePath);
            reader = new BufferedReader(new FileReader(file));
            int line = 1;
            // 一次读入一行，直到读入null为文件结束
            String tempString = null;
            while ((tempString = reader.readLine()) != null) {
                if (line == 1) {
                    strings.add(0, tempString);
                } else if (line == 2) {
                    strings.add(1, tempString);
                } else if (line == 3) {
                    strings.add(2, tempString);
                }
                line++;
            }
            reader.close();
            return strings;
        } catch (Exception e) {
            e.printStackTrace();
            return new ArrayList<String>();
        }
    }
}
