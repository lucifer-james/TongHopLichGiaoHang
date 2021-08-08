package com.example.lucifer.mypoidemo;

import android.util.Log;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Lucifer on 8/1/2021.
 */

public class ThreadXuLyFile extends Thread {

    List<XSSFSheet> lst_ws;
    int[][] arry_Type_BTP = new int[6][4];
    String[] criteria = new String[]{"BTP GÓI 3IN1", "BTP GÓI SOUP", "BTP GÓI RAU","BTP GÓI DẦU","BTP GÓI EKITAI","TỔNG"};

    public ThreadXuLyFile(List<XSSFSheet> lst_ws){
        this.lst_ws = lst_ws;
    }

    public int getRowTypeBTP(String loaiBTP){
        int result=0;
        for (int i=0; i<criteria.length;i++){
            if (loaiBTP.equalsIgnoreCase(criteria[i])){
                result = arry_Type_BTP[i][0];
                break;
            }
        }
        return result;
    }


    @Override
    public void run() {
        super.run();
        int end_row = lst_ws.get(0).getLastRowNum();
        int start_row = 10;


        for (int i= start_row; i< end_row;i++){
            XSSFCell cell_file_goc = lst_ws.get(0).getRow(i).getCell(2);
            cell_file_goc.setCellValue("");
        }

        int m = 0;

        int bat_dau = 10;

        // lấy ra dòng loại BTP của cả 4 sheet và lưu vào mảng 2 chiều arry_Type_BTP
        // Sheet gốc - Sheet 1 - Sheet 2 - Sheet 3
        //     10        10        10        10            tương ứng với dòng của loại BTP là BTP GÓI 3IN1
        //     21        21        21        21            tương ứng với dòng của loại BTP là BTP GÓI SOUP
        // criteria.length = 6 phần tử <=> chỉ số từ 0 - 5
        while (m < criteria.length){
            for (int j=0;j<lst_ws.size();j++){
                // lấy ra dòng của loại BTP, bắt đầu duyệt từ dòng bat_dau
                // nghĩa là lấy dữ liệu cột của mảng arry_Type_BTP
                arry_Type_BTP[m][j] = get_row_of_Type_BTP(lst_ws.get(j),criteria[m],bat_dau);
            }
            // sau khi lấy ra xong thì gán lại cho dòng bắt đầu duyệt loại BTP tiếp theo chính là từ dòng của loại BTP mới lấy ra
            bat_dau = arry_Type_BTP[m][0];
            m++; // tăng m lên 1 đơn vị, nghĩa là đi tới dòng số 2 của mảng arry_Type_BTP

            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        String temp = "";

        for (int z = 0;  z < criteria.length; z++){
            temp+= arry_Type_BTP[z][0] + " - ";

        }

        Log.d("arry_Type_BTP",temp);

        // duyệt qua 6 dòng chứa dòng của loại BTPP trong mảng arry_Type_BTP
        for (int i = 0; i < criteria.length ;i++){
            // lấy dòng chứa loại BTP (từng loại) của file gốc (chỉ số cột bao giờ cũng là 0)
            int file_goc = arry_Type_BTP[i][0];

            // tạo 1 mảng phụ 3 phần tử <=> chứa 3 phần tử số dòng của 3 sheet còn lại
            int[] arry_temp = new int[3];
            for (int j=1; j <4 ; j++){
                // duyệt từ cột số 1 tới cột số 3 của mảng arry_Type_BTP
                // lấy ra phần tử và gán vào mảng arry_temp
                arry_temp[j-1] = arry_Type_BTP[i][j];
            }

            // sau khi có mảng arry_temp thì tìm số lớn nhất của mảng
            int max = find_max_value(arry_temp);

            // sau đó so sánh với số của sheet gốc
            // nếu sheet gốc nhỏ hơn nghĩa là có 1 sheet nào đó đã chèn thêm dòng, làm cho số TT của loại BTP bị thay đổi so với sheet gốc
            if (file_goc<max){
                // lấy ra chênh lệch
                int chenh_lech = max - file_goc;

                // copy dòng và gán lại chỉ số dòng mới cho phần tử thứ i của mảng arry_Type_BTP
                lst_ws.get(0).copyRows(file_goc - 1, file_goc  - chenh_lech,
                        file_goc - chenh_lech - 1, new CellCopyPolicy());
                arry_Type_BTP[i][0] += chenh_lech;

            }
        }

        m=0;

        bat_dau=10;

        while (m < criteria.length - 1 ){
            int k=0;
            start_row = arry_Type_BTP[m][0];
            end_row = arry_Type_BTP[m+1][0];

            List<Object> BTP = new ArrayList<Object>();

            for(int j=1;j < lst_ws.size();j++){
                for (int i=start_row+1;i<end_row;i++){
                    XSSFCell cell = lst_ws.get(j).getRow(i).getCell(2);
                    Object BTP_Name = getCellValue(cell,cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator());
                    if (!BTP.contains(BTP_Name) && !BTP_Name.equals("") && !BTP_Name.equals(0)){
                        BTP.add(BTP_Name.toString());
                    }

                }
                Log.d("BTP Name",BTP.toString());
            }

            for (int i = start_row + 1; i<end_row; i++){
                XSSFCell cell=lst_ws.get(0).getRow(i).getCell(2);
                if (cell==null){
                    lst_ws.get(0).getRow(i).createCell(2);
                }

                if (k < BTP.size()){
                    cell.setCellValue(BTP.get(k).toString());
                    k++;
                }else{
                    break;
                }

            }

            m++;

            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }

    private int get_row_of_Type_BTP(XSSFSheet sheet,String criteria, int from){
        int result = 0;

        int end_row = sheet.getLastRowNum();

        int start_row = from;

        for(int i=start_row;i<=end_row;i++){
            XSSFCell cell = sheet.getRow(i).getCell(1);
            if (cell!=null){
                Object cell_value = getCellValue(cell,cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator());
                if (!cell_value.equals("") && !cell_value.equals(0)){
                    if(cell_value.equals(criteria)){
                        result = i;
                        break;
                    }
                }
            }

        }

        return result;
    }

    private int find_max_value(int[] arry){

        int max =arry[0];

        for(int i=0;i<arry.length;i++){
            if(arry[i]>max){
                max = arry[i];
            }
        }
        return max;
    }

    private Object getCellValue(XSSFCell cell, FormulaEvaluator evaluator){

        Object cellValue = evaluator.evaluate(cell);

        switch (cell.getCellTypeEnum()) {
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case BLANK:
                cellValue = 0;
                break;
            case _NONE:
                cellValue = 0;
                break;
            case ERROR:
                cellValue = 0;
                break;
            case FORMULA:
                cellValue = evaluator.evaluate(cell).getNumberValue();
                break;
            default:
                return 0;
        }

        return cellValue;
    }
}
